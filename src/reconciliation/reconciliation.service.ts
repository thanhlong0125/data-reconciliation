import {
  BadRequestException,
  Injectable,
} from '@nestjs/common';
import ExcelJS from 'exceljs';
import moment from 'moment-timezone';
import { Readable } from 'node:stream';
import { MongoLogService } from '../db/mongo-log.service';
import { UploadedExcelFile } from '../public/uploaded-excel-file.type';
import { B2StorageService } from './B2Storage.service';

export type UploadFileType = 'partner' | 'afc';

export interface ReconciliationColumns {
  key: string;
  time: string;
  price: string;
}

const FILE_TIMEZONE = 'Asia/Ho_Chi_Minh';
const CALENDAR_DATE_FORMAT = 'YYYY-MM-DD';

const CODE = [
  { code: 1, name: 'App - Thẻ quốc thế (MPGS)'},
  { code: 2, name: 'App - TKTT Sacombank' },
  { code: 3, name: 'App - Momo' },
  { code: 4, name: 'App - Zalopay' },
  { code: 5, name: 'App - ShopeePay' },
  // { code: 6, name: 'Tap&go - ZaloPay' },
  // { code: 7, name: 'Tap&go - Momo' },
  // { code: 8, name: 'Tap&go - ShopeePay' },
  { code: 9, name: 'Kiosk - TTT' },
  { code: 10, name: 'Kiosk - ZaloPay' },
  { code: 11, name: 'Kiosk - ShopeePay' },
  { code: 12, name: 'Kiosk - Momo' },
  { code: 13, name: 'Kiosk - Techcombank' },
]

const ENGLISH_MONTH_NUMBERS: Record<string, string> = {
  jan: '01',
  january: '01',
  feb: '02',
  february: '02',
  mar: '03',
  march: '03',
  apr: '04',
  april: '04',
  may: '05',
  jun: '06',
  june: '06',
  jul: '07',
  july: '07',
  aug: '08',
  august: '08',
  sep: '09',
  september: '09',
  oct: '10',
  october: '10',
  nov: '11',
  november: '11',
  dec: '12',
  december: '12',
};
interface CutKeyConfig {
  key: string;
  index: number;
}

interface ReconciliationOptions {
  addKeyAFC?: string;
  addKeyPartner?: string;
  cutKeyAFC?: CutKeyConfig;
  cutKeyPartner?: CutKeyConfig;
  ignoreFailedPartnerTransactions?: boolean;
}

const RECONCILIATION_OPTIONS: Partial<
  Record<number, ReconciliationOptions>
> = {
  2: { cutKeyPartner: { key: '[PCA HCMC Metro HURC]', index: 1 } },
  10: { ignoreFailedPartnerTransactions: true },
  11: { addKeyAFC: 'O' },
  12: { addKeyPartner: 'O' },
  13: { addKeyAFC: 'O' },
};

interface DataRow {
  rowNumber: number;
  key: string;
  displayKey: string;
  price: number;
  values: ExcelJS.CellValue[];
  fullValues: ExcelJS.CellValue[];
}

interface DataSet {
  headers: string[];
  fullHeaders: string[];
  rows: DataRow[];
  refundRows: DataRow[];
  dateColumn: string;
}

export interface DateRange {
  fromDate: string;
  toDate: string;
}

@Injectable()
export class ReconciliationService {
  constructor(
    private readonly b2StorageService: B2StorageService,
    private readonly mongoLogService: MongoLogService,
  ) {}

  async storeUploadedFile(
    file: UploadedExcelFile,
    userKey?: string,
  ): Promise<{
    id: string;
    filename: string;
    key: string;
    links: string[];
    size: number;
    contentType: string;
  }> {
    const filename = file.originalname
      .replace(/\\/g, '/')
      .split('/')
      .pop()
      ?.replace(/[<>:"|?*\x00-\x1f]/g, '_')
      .replace(/[. ]+$/, '')
      .trim();
    if (!filename || !filename.toLowerCase().endsWith('.xlsx')) {
      throw new BadRequestException('Chỉ nhận file Excel .xlsx.');
    }

    const workbook = new ExcelJS.Workbook();
    try {
      const xlsxBuffer = new ArrayBuffer(file.buffer.byteLength);
      new Uint8Array(xlsxBuffer).set(file.buffer);
      await workbook.xlsx.load(xlsxBuffer);
    } catch {
      throw new BadRequestException('File Excel không hợp lệ hoặc bị lỗi.');
    }
    if (workbook.worksheets.length !== 1) {
      throw new BadRequestException(
        'File Excel phải có đúng một sheet dữ liệu.',
      );
    }

    // const outputDirectory = resolve(process.cwd(), 'uploads');
    // const metadataDirectory = resolve(outputDirectory, '.metadata');
    // await mkdir(metadataDirectory, { recursive: true });
    // const outputPath = resolve(outputDirectory, filename);
    // const metadataPath = resolve(metadataDirectory, `${filename}.json`);
    // await writeFile(outputPath, file.buffer, { flag: 'wx' });
    // await writeFile(metadataPath, JSON.stringify({ type, filename }), {
    //   flag: 'wx',
    // });
    const storedFile = await this.b2StorageService.uploadFile({
      originalname: filename,
      mimetype:
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      size: file.buffer.byteLength,
      buffer: file.buffer,
    });
    let id: string;
    try {
      id = await this.mongoLogService.recordUploadedFile({
        name: filename,
        links: [storedFile.url],
        status: 'uploaded',
        key: storedFile.key,
        userKey,
        size: storedFile.size,
        contentType: storedFile.contentType,
      });
    } catch (databaseError) {
      try {
        await this.b2StorageService.deleteFile(storedFile.key);
      } catch (cleanupError) {
        throw new AggregateError(
          [databaseError, cleanupError],
          'Không lưu được metadata MongoDB và không xóa được file trên B2.',
        );
      }
      throw databaseError;
    }
    return {
      id,
      filename,
      key: storedFile.key,
      links: [storedFile.url],
      size: storedFile.size,
      contentType: storedFile.contentType,
    };
  }

  getFile(id: string): Promise<{
    key: string;
    originalName: string;
    mimeType: string;
  }> {
    return this.mongoLogService.getUploadedFile(id);
  }

  getCodes(): { code: number; name: string }[] {
    return CODE.map(({ code, name }) => ({ code, name }));
  }

  async checkCodeAndHandle(
    afcFile: UploadedExcelFile,
    partnerFile: UploadedExcelFile,
    code: number,
    dateRange: DateRange,
    partnerColumns?: ReconciliationColumns,
    afcColumns?: ReconciliationColumns,
    keyRefund?: string,
  ): Promise<Buffer> {
    const selectedCode = CODE.find(item => item.code === code);
    if (!selectedCode) {
      throw new BadRequestException(`Mã code ${code} không hợp lệ.`);
    }
    if (!dateRange) {
      throw new BadRequestException('Vui lòng chọn khoảng ngày đối soát.');
    }
    if (!partnerColumns || !afcColumns) {
      throw new BadRequestException(
        'Vui lòng truyền đủ cột đối soát cho partner và AFC.',
      );
    }
    return this.reconcileFiles(
      afcFile,
      partnerFile,
      dateRange,
      partnerColumns,
      afcColumns,
      keyRefund,
      selectedCode.code,
    );
  }

  private filterByDateRange(
    data: DataSet,
    dateColumnIndex: number,
    dateColumn: string,
    dateRange: DateRange,
    fileLabel: string,
  ): DataSet {
    const resolvedDateColumn = data.dateColumn ?? dateColumn;
    const rows = data.rows.filter((row) => {
      const value = row.values[dateColumnIndex];
      if (value === null || value === undefined || value === '') {
        throw new BadRequestException(
          `Thiếu ngày ở cột "${resolvedDateColumn}", dòng ${row.rowNumber} của ${fileLabel}.`,
        );
      }

      const rowDate = this.toCalendarDate(value);
      if (!rowDate) {
        throw new BadRequestException(
          `Ngày không hợp lệ ở cột "${resolvedDateColumn}", dòng ${row.rowNumber} của ${fileLabel}.`,
        );
      }

      return rowDate >= dateRange.fromDate && rowDate <= dateRange.toDate;
    });

    return { ...data, rows };
  }

  private readReconciliationRows(
    file: UploadedExcelFile,
    fileLabel: string,
    columns: ReconciliationColumns,
    keyRefund?: string,
  ): Promise<DataSet> {
    const selectedColumnsForInput = [
      columns.key,
      columns.price,
      columns.time,
      ...(keyRefund ? [keyRefund] : []),
    ];

    return this.readRows(
      file,
      fileLabel,
      selectedColumnsForInput,
      columns.key,
      columns.price,
      columns.time,
      keyRefund,
    );
  }

  private toCalendarDate(value: ExcelJS.CellValue): string | undefined {
    if (value instanceof Date) {
      return Number.isNaN(value.getTime())
        ? undefined
        : moment(value).tz(FILE_TIMEZONE).format(CALENDAR_DATE_FORMAT);
    }

    let text: string;
    if (typeof value === 'number') {
      if (
        Number.isInteger(value) &&
        value >= 10000101 &&
        value <= 99991231
      ) {
        text = String(value);
      } else if (value > 0) {
        const excelDate = new Date(
          Date.UTC(1899, 11, 30) + Math.round(value * 86400000),
        );
        return Number.isNaN(excelDate.getTime())
          ? undefined
          : moment
              .utc(excelDate)
              .tz(FILE_TIMEZONE)
              .format(CALENDAR_DATE_FORMAT);
      } else {
        return undefined;
      }
    } else if (typeof value === 'string') {
      text = value.trim();
    } else {
      return undefined;
    }

    let year: string;
    let month: string;
    let day: string;
    let hourText: string | undefined;
    let minuteText: string | undefined;
    let secondText: string | undefined;
    let millisecondText: string | undefined;
    let meridiem: string | undefined;
    let timezone: string | undefined;
    const compactDate = /^(\d{4})(\d{2})(\d{2})$/.exec(text);
    const numericDateTime =
      /^(\d{1,4})[./-](\d{1,2})[./-](\d{1,4})(?:[ T,]+(\d{1,2}):(\d{2})(?::(\d{2})(?:[.,](\d{1,3}))?)?\s*(AM|PM)?\s*(Z|[+-]\d{2}:?\d{2})?)?$/i.exec(
        text,
      );
    if (compactDate) {
      [, year, month, day] = compactDate;
    } else if (numericDateTime) {
      const [
        ,
        first,
        secondPart,
        third,
        parsedHour,
        parsedMinute,
        parsedSecond,
        parsedMillisecond,
        parsedMeridiem,
        parsedTimezone,
      ] = numericDateTime;
      if (first.length === 4) {
        [year, month, day] = [first, secondPart, third];
      } else if (third.length === 4) {
        year = third;
        if (Number(first) > 12) {
          [day, month] = [first, secondPart];
        } else if (Number(secondPart) > 12) {
          [month, day] = [first, secondPart];
        } else {
          [day, month] = [first, secondPart];
        }
      } else {
        return undefined;
      }
      hourText = parsedHour;
      minuteText = parsedMinute;
      secondText = parsedSecond;
      millisecondText = parsedMillisecond;
      meridiem = parsedMeridiem;
      timezone = parsedTimezone;
    } else {
      const dayFirst =
        /^(\d{1,2})\s+([a-z]+)\s+(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/i.exec(
          text,
        );
      const monthFirst =
        /^([a-z]+)\s+(\d{1,2}),?\s+(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/i.exec(
          text,
        );
      if (dayFirst) {
        [, day, month, year, hourText, minuteText, secondText] = dayFirst;
      } else if (monthFirst) {
        [, month, day, year, hourText, minuteText, secondText] = monthFirst;
      } else {
        return undefined;
      }
      month = ENGLISH_MONTH_NUMBERS[month.toLowerCase()];
      if (!month) {
        return undefined;
      }
    }

    const numericYear = Number(year);
    const numericMonth = Number(month);
    const numericDay = Number(day);
    const calendarDate = new Date(0);
    calendarDate.setUTCHours(0, 0, 0, 0);
    calendarDate.setUTCFullYear(
      numericYear,
      numericMonth - 1,
      numericDay,
    );
    if (
      calendarDate.getUTCFullYear() !== numericYear ||
      calendarDate.getUTCMonth() !== numericMonth - 1 ||
      calendarDate.getUTCDate() !== numericDay
    ) {
      return undefined;
    }

    const dateText = `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
    if (hourText === undefined) {
      return dateText;
    }

    let hour = Number(hourText);
    const minute = Number(minuteText);
    const second = Number(secondText ?? 0);
    if (!Number.isInteger(minute)) {
      return undefined;
    }
    if (meridiem) {
      if (hour < 1 || hour > 12) {
        return undefined;
      }
      hour = (hour % 12) + (meridiem.toUpperCase() === 'PM' ? 12 : 0);
    } else if (hour > 23) {
      return undefined;
    }
    if (minute > 59 || second > 59) {
      return undefined;
    }

    const time = `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}:${String(second).padStart(2, '0')}.${(millisecondText ?? '0').padEnd(3, '0')}`;
    if (!timezone) {
      const timestamp = new Date(calendarDate.getTime());
      timestamp.setUTCHours(
        hour,
        minute,
        second,
        Number((millisecondText ?? '0').padEnd(3, '0')),
      );
      return moment
        .utc(timestamp)
        .tz(FILE_TIMEZONE)
        .format(CALENDAR_DATE_FORMAT);
    }

    const normalizedTimezone =
      timezone.toUpperCase() === 'Z'
        ? 'Z'
        : timezone.includes(':')
          ? timezone
          : `${timezone.slice(0, 3)}:${timezone.slice(3)}`;
    if (normalizedTimezone !== 'Z') {
      const offset = /^([+-])(\d{2}):(\d{2})$/.exec(normalizedTimezone);
      if (!offset || Number(offset[2]) > 23 || Number(offset[3]) > 59) {
        return undefined;
      }
    }

    const timestamp = new Date(`${dateText}T${time}${normalizedTimezone}`);
    return Number.isNaN(timestamp.getTime())
      ? undefined
      : moment(timestamp).tz(FILE_TIMEZONE).format(CALENDAR_DATE_FORMAT);
  }

  private async refund(){
    
  }

  private async reconcileFiles(
    afcFile: UploadedExcelFile,
    partnerFile: UploadedExcelFile,
    dateRange: DateRange,
    partnerColumns: ReconciliationColumns,
    afcColumns: ReconciliationColumns,
    keyRefund: string | undefined,
    code: number,
  ): Promise<Buffer> {
    const options = RECONCILIATION_OPTIONS[code] ?? {};
    const partner = await this.readReconciliationRows(
      partnerFile,
      'partnerFile',
      partnerColumns,
      keyRefund,
    );
    const afc = await this.readReconciliationRows(
      afcFile,
      'afcFile',
      afcColumns,
      keyRefund,
    );
    const partnerRows = options.ignoreFailedPartnerTransactions
      ? {
          ...partner,
          rows: partner.rows.filter((row) => {
            const status = row.values[3];
            return !(
              typeof status === 'string' &&
              status.trim().toLocaleLowerCase('vi-VN').includes('thất bại')
            );
          }),
        }
      : partner;

    return this.recociliation(
      this.filterByDateRange(
        partnerRows,
        2,
        partnerColumns.time,
        dateRange,
        'partnerFile',
      ),
      this.filterByDateRange(
        afc,
        2,
        afcColumns.time,
        dateRange,
        'afcFile',
      ),
      options.addKeyAFC,
      options.addKeyPartner,
      options.cutKeyAFC,
      options.cutKeyPartner,
    );
  }
  private async recociliation(
    partnerData: DataSet,
    afcData: DataSet,
    addKeyAFC?: string,
    addKeyPartner?: string,
    cutKeyAFC ?: CutKeyConfig,
    cutKeyPartner?: CutKeyConfig,
  ): Promise<Buffer> {

    const partnerRowsByMatch = this.findUnmatchedRows(
      partnerData.rows,
      afcData.rows,
      addKeyPartner,
      addKeyAFC,
      cutKeyPartner,
      cutKeyAFC,
    );
    const afcRowsByMatch = this.findUnmatchedRows(
      afcData.rows,
      partnerData.rows,
      addKeyAFC,
      addKeyPartner,
      cutKeyAFC,
      cutKeyPartner,
    );
    const partnerThuDuRows = partnerRowsByMatch.unmatchedRows;
    const partnerNhieuHonAfcRows = partnerRowsByMatch.extraRows;
    const afcThuDuRows = afcRowsByMatch.unmatchedRows;
    const afcNhieuHonDoiTacRows = afcRowsByMatch.extraRows;

    const result = new ExcelJS.Workbook();
    const tongHopGiaiTrinh = result.addWorksheet('TongHopGiaiTrinh');
    tongHopGiaiTrinh.addRow([
      'STT',
      'Giải trình',
      'Chi tiết giao dịch tại',
    ]);
    tongHopGiaiTrinh.addRow([
      1,
      `AFC và Đối tác đều có giao dịch; AFC thu dư ${afcThuDuRows.length} lệnh, tổng tiền ${afcThuDuRows.reduce((total, row) => total + row.price, 0)} VND`,
      'AFC thu dư',
    ]);
    tongHopGiaiTrinh.addRow([
      2,
      `Đối tác và AFC đều có giao dịch; Đối tác thu dư ${partnerThuDuRows.length} lệnh, tổng tiền ${partnerThuDuRows.reduce((total, row) => total + row.price, 0)} VND`,
      'Đối tác thu dư',
    ]);
    tongHopGiaiTrinh.addRow([
      3,
      `AFC có ${afcNhieuHonDoiTacRows.length} giao dịch trị giá ${afcNhieuHonDoiTacRows.reduce((total, row) => total + row.price, 0)} VND nhưng Đối tác không có giao dịch tương ứng`,
      'AFC nhiều hơn đối tác',
    ]);
    tongHopGiaiTrinh.addRow([
      4,
      `Đối tác có ${partnerNhieuHonAfcRows.length} giao dịch trị giá ${partnerNhieuHonAfcRows.reduce((total, row) => total + row.price, 0)} VND nhưng AFC không có giao dịch tương ứng`,
      'Đối tác nhiều hơn AFC',
    ]);

    const afcThuDu = result.addWorksheet('AFC thu dư');
    afcThuDu.addRow(afcData.fullHeaders);
    for (const row of afcThuDuRows) {
      afcThuDu.addRow(row.fullValues);
    }

    const doiTacThuDu = result.addWorksheet('Đối tác thu dư');
    doiTacThuDu.addRow(partnerData.fullHeaders);
    for (const row of partnerThuDuRows) {
      doiTacThuDu.addRow(row.fullValues);
    }

    const afcNhieuHonDoiTac = result.addWorksheet('AFC nhiều hơn đối tác');
    afcNhieuHonDoiTac.addRow(afcData.fullHeaders);
    for (const row of afcNhieuHonDoiTacRows) {
      afcNhieuHonDoiTac.addRow(row.fullValues);
    }

    const doiTacNhieuHonAfc = result.addWorksheet('Đối tác nhiều hơn AFC');
    doiTacNhieuHonAfc.addRow(partnerData.fullHeaders);
    for (const row of partnerNhieuHonAfcRows) {
      doiTacNhieuHonAfc.addRow(row.fullValues);
    }

    const refundEntries = [
      ...partnerData.refundRows.map((row) => ({
        source: 'partnerFile',
        headers: partnerData.fullHeaders,
        row,
      })),
      ...afcData.refundRows.map((row) => ({
        source: 'afcFile',
        headers: afcData.fullHeaders,
        row,
      })),
    ];
    if (refundEntries.length > 0) {
      const refundHeaders = Array.from(
        new Set(refundEntries.flatMap((entry) => entry.headers)),
      );
      const hoanTien = result.addWorksheet('Hoàn tiền');
      hoanTien.addRow(['Nguồn', ...refundHeaders]);

      for (const entry of refundEntries) {
        const valuesByHeader = new Map(
          entry.headers.map((header, index) => [
            header,
            entry.row.fullValues[index],
          ]),
        );
        hoanTien.addRow([
          entry.source,
          ...refundHeaders.map((header) => valuesByHeader.get(header)),
        ]);
      }
    }

    return Buffer.from(await result.xlsx.writeBuffer());
  }

  private findUnmatchedRows(
    sourceRows: DataRow[],
    counterpartRows: DataRow[],
    sourceAddKey?: string,
    counterpartAddKey?: string,
    sourceCutKeyAtSpace?: CutKeyConfig,
    counterpartCutKeyAtSpace?: CutKeyConfig,
  ): {
    unmatchedRows: DataRow[];
    extraRows: DataRow[];
  } {
    const addPrefix = (key: string, prefix?: string): string => {
      const normalizedPrefix = prefix?.toLocaleLowerCase('vi-VN')
      if (!normalizedPrefix || key.startsWith(normalizedPrefix)) {
        return key;
      }
      return `${normalizedPrefix}${key}`;
    };
    const normalizeKey = (key: string, config?: CutKeyConfig): string => {
      if (!config) {
        return key.trim();
      }
      return key.split(config.key)[config.index]?.trim() ?? '';
    };

    const counterpartCountsByKey = new Map<string, number>();
    for (const row of counterpartRows) {
      if (!row.key) {
        continue;
      }
      const key = addPrefix(
        normalizeKey(row.key, counterpartCutKeyAtSpace),
        counterpartAddKey,
      );
      counterpartCountsByKey.set(
        key,
        (counterpartCountsByKey.get(key) ?? 0) + 1,
      );

    }

    const sourceOccurrencesByKey = new Map<string, number>();
    const unmatchedRows: DataRow[] = [];
    const extraRows: DataRow[] = [];
    for (const row of sourceRows) {
      if (!row.key) {
        extraRows.push(row);
        continue;
      }

      const key = addPrefix(
        normalizeKey(row.key, sourceCutKeyAtSpace),
        sourceAddKey,
      );
      const sourceOccurrence =
        (sourceOccurrencesByKey.get(key) ?? 0) + 1;
      sourceOccurrencesByKey.set(key, sourceOccurrence);

      const counterpartCount = counterpartCountsByKey.get(key) ?? 0;
      if (counterpartCount === 0) {
        extraRows.push(row);
      } else if (sourceOccurrence > counterpartCount) {
        unmatchedRows.push(row);
      }
    }

    return { unmatchedRows, extraRows };
  }

  private async readRows(
    file: UploadedExcelFile,
    fileLabel: string,
    selectedColumns: string[],
    keyColumn: string,
    priceColumn: string,
    dateColumn: string,
    refundStatusColumn?: string,
  ): Promise<DataSet> {
    if (!file.originalname.toLowerCase().endsWith('.xlsx')) {
      throw new BadRequestException(`${fileLabel} phải là file Excel .xlsx.`);
    }

    const rows: DataRow[] = [];
    const refundRows: DataRow[] = [];
    let firstSheetFound = false;
    let headersFound = false;
    let fullHeaders: string[] = [];

    try {
      const reader = new ExcelJS.stream.xlsx.WorkbookReader(
        Readable.from([file.buffer]),
        {
          worksheets: 'emit',
          sharedStrings: 'cache',
          styles: 'cache',
          hyperlinks: 'ignore',
          entries: 'ignore',
        },
      );

      for await (const worksheet of reader) {
        firstSheetFound = true;
        let selectedColumnIndexes: number[] = [];
        let keyColumnIndex = 0;
        let priceColumnIndex = 0;
        let refundStatusColumnIndex: number | undefined;
        let fullHeaderCount = 0;

        for await (const row of worksheet) {
          if (row.number === 1) {
            const headerIndexes = new Map<string, number>();
            row.eachCell((cell, columnNumber) => {
              const normalizedHeader = this.normalizeHeader(cell.text);
              if (!headerIndexes.has(normalizedHeader)) {
                headerIndexes.set(normalizedHeader, columnNumber);
              }
            });
            const lastColumnNumber = headerIndexes.size
              ? Math.max(...Array.from(headerIndexes.values()))
              : 0;
            fullHeaderCount = lastColumnNumber;
            fullHeaders = Array.from(
              { length: fullHeaderCount },
              (_, index) => row.getCell(index + 1).text,
            );

            const findColumn = (headerName: string): number => {
              const columnIndex = headerIndexes.get(
                this.normalizeHeader(headerName),
              );
              if (columnIndex === undefined) {
                throw new BadRequestException(
                  `Không tìm thấy cột "${headerName}" trong dòng tiêu đề của ${fileLabel}.`,
                );
              }
              return columnIndex;
            };

            selectedColumnIndexes = selectedColumns.map(findColumn);
            keyColumnIndex = findColumn(keyColumn);
            priceColumnIndex = findColumn(priceColumn);
            refundStatusColumnIndex = refundStatusColumn
              ? findColumn(refundStatusColumn)
              : undefined;
            headersFound = true;
            continue;
          }

          if (!headersFound) {
            continue;
          }
          if (!row.hasValues) {
            continue;
          }
          const dataRow = this.toDataRow(
            row,
            keyColumnIndex,
            priceColumnIndex,
            selectedColumnIndexes,
            fullHeaderCount,
          );
          const isRefund =
            refundStatusColumnIndex !== undefined &&
            ['hoàn tiền', 'hoàn trả', '0'].includes(
              row
                .getCell(refundStatusColumnIndex)
                .text.trim()
                .toLocaleLowerCase('vi-VN'),
            );

          (isRefund ? refundRows : rows).push(dataRow);
        }

        return {
          headers: [...selectedColumns],
          fullHeaders,
          rows,
          refundRows,
          dateColumn,
        };
      }
    } catch (error) {
      if (error instanceof BadRequestException) {
        throw error;
      }
      throw new BadRequestException(`Không thể đọc file ${fileLabel}.`);
    }

    if (!firstSheetFound) {
      throw new BadRequestException(`File ${fileLabel} không có sheet dữ liệu.`);
    }
    if (!headersFound) {
      throw new BadRequestException(
        `Không tìm thấy dòng tiêu đề trong file ${fileLabel}.`,
      );
    }

    return {
      headers: [...selectedColumns],
      fullHeaders,
      rows,
      refundRows,
      dateColumn,
    };
  }

  private normalizeHeader(header: string): string {
    return header.trim().normalize('NFC').toLocaleLowerCase('vi-VN');
  }

  private toDataRow(
    row: ExcelJS.Row,
    keyColumn: number,
    priceColumn: number,
    selectedColumnIndexes: number[],
    fullHeaderCount = 0,
  ): DataRow {
    const keyCell = row.getCell(keyColumn);
    const displayKey = this.readKeyAsString(keyCell.value, keyCell.text);
    const priceCell = row.getCell(priceColumn);
    return {
      rowNumber: row.number,
      key: displayKey.toLocaleLowerCase('vi-VN'),
      displayKey,
      price: this.parseAmount(priceCell.value, priceCell.text),
      values: selectedColumnIndexes.map(
        (columnIndex) => row.getCell(columnIndex).value,
      ),
      fullValues: Array.from(
        { length: fullHeaderCount },
        (_, index) => row.getCell(index + 1).value,
      ),
    };
  }

  private readKeyAsString(value: ExcelJS.CellValue, text: string): string {
    if (typeof value === 'string') {
      return value.trim();
    }
    if (typeof value === 'number' && Number.isFinite(value)) {
      return this.numberToPlainString(value).replace(/\.0+$/, '').trim();
    }
    return text.trim();
  }

  private numberToPlainString(value: number): string {
    const valueText = value.toString();
    if (!/[eE]/.test(valueText)) {
      return valueText;
    }

    const [mantissa, exponentText] = valueText.toLowerCase().split('e');
    const exponent = Number(exponentText);
    const negative = mantissa.startsWith('-');
    const unsignedMantissa = negative ? mantissa.slice(1) : mantissa;
    const decimalIndex = unsignedMantissa.indexOf('.');
    const integerDigits =
      decimalIndex === -1
        ? unsignedMantissa
        : unsignedMantissa.slice(0, decimalIndex);
    const fractionDigits =
      decimalIndex === -1
        ? ''
        : unsignedMantissa.slice(decimalIndex + 1);
    const digits = integerDigits + fractionDigits;
    const shiftedDecimalIndex = integerDigits.length + exponent;
    const sign = negative ? '-' : '';

    if (shiftedDecimalIndex <= 0) {
      return `${sign}0.${'0'.repeat(-shiftedDecimalIndex)}${digits}`;
    }
    if (shiftedDecimalIndex >= digits.length) {
      return `${sign}${digits}${'0'.repeat(shiftedDecimalIndex - digits.length)}`;
    }
    return `${sign}${digits.slice(0, shiftedDecimalIndex)}.${digits.slice(shiftedDecimalIndex)}`;
  }

  private parseAmount(value: ExcelJS.CellValue, text: string): number {
    if (typeof value === 'number') {
      return Number.isFinite(value) ? Math.round(value) : 0;
    }

    const normalized = text
      .trim()
      .replace(/\s+/g, '')
      .replace(/[^\d.,+-]/g, '');
    if (!normalized) {
      return 0;
    }

    const negative = normalized.startsWith('-');
    const unsignedAmount = normalized.replace(/^[+-]/, '');
    const decimalPart = /\.(\d{2})$/.exec(unsignedAmount)?.[1];
    const integerPart = decimalPart
      ? unsignedAmount.slice(0, -3)
      : unsignedAmount;
    const integerDigits = integerPart.replace(/[.,]/g, '');
    if (!/^\d+$/.test(integerDigits)) {
      return 0;
    }

    const amount = Number(
      `${negative ? '-' : ''}${integerDigits}${decimalPart ? `.${decimalPart}` : ''}`,
    );
    return Number.isFinite(amount) ? Math.round(amount) : 0;
  }

}
