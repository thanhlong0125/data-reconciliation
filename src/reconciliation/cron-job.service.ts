import {
  Injectable,
  Logger,
  OnModuleDestroy,
  OnModuleInit,
} from '@nestjs/common';
import { mkdir, writeFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { MongoLogService } from '../db/mongo-log.service';
import { ReconciliationService } from './reconciliation.service';
import { UploadedExcelFile } from './uploaded-excel-file.type';

const MAX_FILE_SIZE = 20 * 1024 * 1024;
const JOB_INTERVAL = 10 * 60 * 1000;

interface SourceFilesResponse {
  partnerFileUrl: string;
  afcFileUrl: string;
}

@Injectable()
export class CronJobService implements OnModuleInit, OnModuleDestroy {
  private readonly logger = new Logger(CronJobService.name);
  private interval: ReturnType<typeof setInterval> | undefined;
  private isRunning = false;

  constructor(
    private readonly reconciliationService: ReconciliationService,
    private readonly mongoLogService: MongoLogService,
  ) {}

  onModuleInit(): void {
    this.interval = setInterval(() => {
      void this.runJob();
    }, JOB_INTERVAL);
  }

  onModuleDestroy(): void {
    if (this.interval) {
      clearInterval(this.interval);
    }
  }

  private async runJob(): Promise<void> {
    if (this.isRunning) {
      this.logger.warn('Đang chạy cron đối soát; bỏ qua lần chạy này.');
      await this.mongoLogService.recordSystemEvent('cron_skipped', {
        reason: 'previous_run_still_running',
      });
      return;
    }

    const sourceUrl = process.env.RECONCILIATION_SOURCE_URL?.trim();
    if (!sourceUrl) {
      this.logger.warn(
        'Thiếu env RECONCILIATION_SOURCE_URL',
      );
      await this.mongoLogService.recordSystemEvent('cron_skipped', {
        reason: 'missing_RECONCILIATION_SOURCE_URL',
      });
      return;
    }

    this.isRunning = true;
    try {
      const code = Number(process.env.RECONCILIATION_CODE);
      const fromDate = process.env.RECONCILIATION_FROM_DATE?.trim() ?? '';
      const toDate = process.env.RECONCILIATION_TO_DATE?.trim() ?? '';
      if (
        !Number.isInteger(code) ||
        !this.reconciliationService
          .getCodes()
          .some((availableCode) => availableCode.code === code) ||
        !this.isValidDate(fromDate) ||
        !this.isValidDate(toDate) ||
        fromDate > toDate
      ) {
        throw new Error(
          'Cấu hình RECONCILIATION_CODE, RECONCILIATION_FROM_DATE hoặc RECONCILIATION_TO_DATE không hợp lệ.',
        );
      }

      const sourceEndpoint = new URL(sourceUrl);
      if (sourceEndpoint.protocol !== 'https:') {
        throw new Error('URL API nguồn phải sử dụng HTTPS.');
      }

      const sourceResponse = await fetch(sourceEndpoint, {
        signal: AbortSignal.timeout(60_000),
        redirect: 'error',
      });
      if (!sourceResponse.ok) {
        throw new Error(
          `API nguồn trả về HTTP ${sourceResponse.status}.`,
        );
      }

      const fileUrls = await sourceResponse.json() as Partial<SourceFilesResponse>;
      if (
        typeof fileUrls.partnerFileUrl !== 'string' ||
        typeof fileUrls.afcFileUrl !== 'string'
      ) {
        throw new Error(
          'Response API nguồn phải có partnerFileUrl và afcFileUrl.',
        );
      }

      const [partnerBuffer, afcBuffer] = await Promise.all([
        this.downloadExcel(fileUrls.partnerFileUrl),
        this.downloadExcel(fileUrls.afcFileUrl),
      ]);
      const output = await this.reconciliationService.checkCodeAndHandle(
        this.toUploadedFile('afc.xlsx', afcBuffer),
        this.toUploadedFile('partner.xlsx', partnerBuffer),
        code,
        { fromDate, toDate },
      );

      const outputDirectory = resolve(
        process.env.RECONCILIATION_OUTPUT_DIR?.trim() || 'output',
      );
      await mkdir(outputDirectory, { recursive: true });
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const outputPath = join(
        outputDirectory,
        `reconciliation-${code}-${timestamp}.xlsx`,
      );
      await writeFile(outputPath, output);
      this.logger.log(`Đã tạo file đối soát: ${outputPath}`);
      await this.mongoLogService.recordSystemEvent('cron_success', {
        code,
        outputPath,
        fromDate,
        toDate,
      });
    } catch (error) {
      const message =
        error instanceof Error ? error.message : 'Lỗi không xác định.';
      this.logger.error(`Cron đối soát thất bại: ${message}`);
      await this.mongoLogService.recordSystemEvent('cron_failure', {
        message,
      });
    } finally {
      this.isRunning = false;
    }
  }

  private async downloadExcel(url: string): Promise<Buffer> {
    const fileUrl = new URL(url);
    if (fileUrl.protocol !== 'https:') {
      throw new Error('URL tải file Excel phải sử dụng HTTPS.');
    }

    const response = await fetch(fileUrl, {
      signal: AbortSignal.timeout(60_000),
      redirect: 'error',
    });
    if (!response.ok) {
      throw new Error(`Tải file Excel thất bại: HTTP ${response.status}.`);
    }

    const contentLength = Number(response.headers.get('content-length'));
    if (contentLength > MAX_FILE_SIZE) {
      throw new Error('File Excel tải về vượt quá giới hạn 20 MB.');
    }
    if (!response.body) {
      throw new Error('Response tải file Excel không có dữ liệu.');
    }

    const reader = response.body.getReader();
    const chunks: Uint8Array[] = [];
    let totalSize = 0;
    while (true) {
      const { done, value } = await reader.read();
      if (done) {
        break;
      }
      totalSize += value.byteLength;
      if (totalSize > MAX_FILE_SIZE) {
        await reader.cancel();
        throw new Error('File Excel tải về vượt quá giới hạn 20 MB.');
      }
      chunks.push(value);
    }

    const buffer = Buffer.concat(chunks);
    if (buffer.length < 4 || buffer.subarray(0, 2).toString() !== 'PK') {
      throw new Error('File tải về không phải file Excel .xlsx hợp lệ.');
    }
    return buffer;
  }

  private toUploadedFile(
    originalname: string,
    buffer: Buffer,
  ): UploadedExcelFile {
    return { originalname, buffer };
  }

  private isValidDate(value: string): boolean {
    const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
    if (!match) {
      return false;
    }

    const [, yearText, monthText, dayText] = match;
    const year = Number(yearText);
    const month = Number(monthText);
    const day = Number(dayText);
    const date = new Date(Date.UTC(year, month - 1, day));
    return (
      date.getUTCFullYear() === year &&
      date.getUTCMonth() === month - 1 &&
      date.getUTCDate() === day
    );
  }
}
