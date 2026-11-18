import {
  BadRequestException,
  Body,
  Controller,
  Post,
  Res,
  UploadedFiles,
  UseInterceptors,
} from '@nestjs/common';
import { FileFieldsInterceptor } from '@nestjs/platform-express';
import { Response } from 'express';
import { ReconciliationService } from './reconciliation.service';
import { UploadedExcelFile } from './uploaded-excel-file.type';

@Controller('reconciliation')
export class ReconciliationController {
  constructor(private readonly reconciliationService: ReconciliationService) {}

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
    @Body() body: { code: number },
    @Res() response: Response,
  ): Promise<void> {
    const partnerFile = files?.partnerFile?.[0];
    const afcFile = files?.afcFile?.[0];
    const code = body?.code;
    if (!partnerFile || !afcFile) {
      throw new BadRequestException(
        'Vui lòng tải lên đủ partnerFile và afcFile.',
      );
    }
    
    const output = await this.reconciliationService.checkCodeAndHandle(
      afcFile,
      partnerFile,
      code,
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
