import {
  BadRequestException,
  Body,
  Controller,
  Get,
  Post,
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
import { MongoLogService } from '../db/mongo-log.service';
import { ReconciliationService } from './reconciliation.service';
import type { DateRange } from './reconciliation.service';
import { UploadedExcelFile } from './uploaded-excel-file.type';

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
  ) {}

  @Get('codes')
  getCodes(): { code: number; name: string }[] {
    return this.reconciliationService.getCodes();
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
  ): Promise<{ type: string; filename: string; path: string }> {
    if (!file) {
      throw new BadRequestException('Vui lòng gửi một file trong field "file".');
    }
    if (body?.type !== 'partner' && body?.type !== 'afc') {
      throw new BadRequestException(
        'Field "type" phải có giá trị "partner" hoặc "afc".',
      );
    }

    const result = await this.reconciliationService.storeUploadedFile(
      file,
      body.type,
    );
    await this.mongoLogService.recordUserAction(
      request.ip,
      'upload_file',
      { type: result.type, filename: result.filename },
    );
    return result;
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
    body: { code?: string; fromDate?: string; toDate?: string },
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
