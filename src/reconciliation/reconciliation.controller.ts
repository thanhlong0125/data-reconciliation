import {
  BadRequestException,
  Body,
  Controller,
  Get,
  Param,
  Post,
  Query,
  Req,
  Res,
  UploadedFile,
  UploadedFiles,
  UseInterceptors,
} from '@nestjs/common';
import {
  FileFieldsInterceptor,
  FileInterceptor,
} from '@nestjs/platform-express';
import type { Request, Response } from 'express';
import { parsePagination } from '../common/pagination';
import { MongoLogService } from '../db/mongo-log.service';
import { B2StorageService } from './B2Storage.service';
import { ReconciliationService } from './reconciliation.service';
import type { DateRange } from './reconciliation.service';
import { UploadedExcelFile } from '../public/uploaded-excel-file.type';

function isValidDate(value: string): boolean {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) {
    return false;
  }

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  return (
    date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day
  );
}

@Controller('reconciliation')
export class ReconciliationController {
  constructor(
    private readonly reconciliationService: ReconciliationService,
    private readonly mongoLogService: MongoLogService,
    private readonly b2StorageService: B2StorageService,
  ) { }

  @Get('codes')
  getCodes(): Promise<Array<{ code: number; name: string }>> {
    return this.reconciliationService.getCodes();
  }

  @Get('codes/:code')
  getCodeConfig(@Param('code') codeValue: string) {
    const code = Number(codeValue);
    if (!Number.isInteger(code)) {
      throw new BadRequestException('Mã code không hợp lệ.');
    }

    return this.reconciliationService.getCodeConfig(code);
  }

  @Post('upload')
  @UseInterceptors(
    FileInterceptor('file', {
      limits: { fileSize: 70 * 1024 * 1024, files: 1 },
    }),
  )
  async uploadFile(
    @UploadedFile() file: UploadedExcelFile | undefined,
    @Body() body: { type?: string },
    @Req() request: Request,
  ): Promise<{
    filename: string;
    id: string;
    key: string;
    links: string[];
    size: number;
    contentType: string;
  }> {
    try {
      if (!file) {
        throw new BadRequestException(
          'Vui lòng gửi một file trong field "file".',
        );
      }
      const result = await this.reconciliationService.storeUploadedFile(
        file,
        request.ip,
      );
      await this.mongoLogService.recordUserAction(
        request.ip,
        'upload_file',
        { filename: result.filename, key: result.key },
      );
      return result;
    } catch (error) {
      await this.mongoLogService.recordSystemEvent('upload_failed', {
        filename: file?.originalname,
        userKey: request.ip,
        message:
          error instanceof Error ? error.message : 'Lỗi không xác định.',
      });
      throw error;
    }
  }

  @Get('download/:id')
  async download(
    @Param('id') id: string,
    @Res() res: Response,
  ) {
    const file = await this.reconciliationService.getFile(id);

    const body = await this.b2StorageService.downloadFile(file.key);

    res.set({
      'Content-Type': file.mimeType,
      'Content-Disposition': `attachment; filename="${encodeURIComponent(file.originalName)}"`,
    });

    body.pipe(res);
  }

  @Get('list-uploaded-files')
  async listUploadedFiles(
    @Req() request: Request,
    @Query('page') page?: string,
    @Query('limit') limit?: string,
  ) {
    return this.mongoLogService.getUploadedFiles(
      request.ip,
      parsePagination(page, limit),
    );
  }

  @Post('readfile')
  @UseInterceptors(
    FileFieldsInterceptor(
      [
        { name: 'partnerFile', maxCount: 1 },
        { name: 'afcFile', maxCount: 1 },
      ],
      { limits: { fileSize: 20 * 1024 * 1024, files: 2 } },
    ),
  )

  async filterRefunds(
    @UploadedFiles()
    files: {
      partnerFile?: UploadedExcelFile[];
      afcFile?: UploadedExcelFile[];
    },
    @Body()
    body: {
      code?: string;
      fromDate?: string;
      toDate?: string;
    },
    @Res() response: Response,
    @Req() request: Request,
  ): Promise<void> {
    const partnerFile = files?.partnerFile?.[0];
    const afcFile = files?.afcFile?.[0];
    const code = Number(body?.code);
    if (!partnerFile || !afcFile) {
      throw new BadRequestException(
        'Vui lòng tải lên đủ partnerFile và afcFile.',
      );
    }

    if (!Number.isInteger(code)) {
      throw new BadRequestException('Vui lòng chọn mã đối soát hợp lệ.');
    }

    const fromDate = body?.fromDate?.trim() ?? '';
    const toDate = body?.toDate?.trim() ?? '';
    if (!isValidDate(fromDate) || !isValidDate(toDate)) {
      throw new BadRequestException(
        'Vui lòng nhập Từ ngày và Đến ngày hợp lệ.',
      );
    }
    if (fromDate > toDate) {
      throw new BadRequestException(
        'Từ ngày không được sau ngày kết thúc.',
      );
    }
    const dateRange: DateRange = { fromDate, toDate };

    const output = await this.reconciliationService.checkCodeAndHandle(
      afcFile,
      partnerFile,
      code,
      dateRange,
    );
    await this.mongoLogService.recordUserAction(
      request.ip,
      'reconciliation',
      {
        code,
        fromDate,
        toDate,
        partnerFilename: partnerFile.originalname,
        afcFilename: afcFile.originalname,
      },
    );
    response
      .set({
        'Content-Type':
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition':
          'attachment; filename="doi-soat-1-1.xlsx"',
      })
      .end(output);
  }
}
