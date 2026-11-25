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

const FILE_TIMEZONE = 'Asia/Ho_Chi_Minh';
const CALENDAR_DATE_FORMAT = 'YYYY-MM-DD';

const CODE = [
  { code: 1, name: 'App - Thẻ quốc thế (MPGS)'},
  { code: 2, name: 'App - TKTT Sacombank' },
  { code: 3, name: 'App - Momo' },
  { code: 4, name: 'App - Zalopay' },
  { code: 5, name: 'App - ShopeePay' },
  { code: 6, name: 'Tap&go - ZaloPay' },
  { code: 7, name: 'Tap&go - Momo' },
  { code: 8, name: 'Tap&go - ShopeePay' },
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

interface DataRow {
  rowNumber: number;
  key: string;
  displayKey: string;
  price: number;
  transactionType?: string;
  values: ExcelJS.CellValue[];
  fullValues: ExcelJS.CellValue[];
}

interface DataSet {
  headers: string[];
  fullHeaders: string[];
  rows: DataRow[];
}

interface RefundData {
  refunds: DataSet;
  remaining: DataSet;
}

interface WrongPriceRow {
  afcRow: DataRow;
  partnerRow: DataRow;
  priceDifference: number;
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
    dateRange?: DateRange,
  ): Promise<Buffer> {
    const selectedCode = CODE.find(item => item.code === code);
    if (!selectedCode) {
      throw new BadRequestException(`Mã code ${code} không hợp lệ.`);
    }
    if (!dateRange) {
      throw new BadRequestException('Vui lòng chọn khoảng ngày đối soát.');
    }
    switch (selectedCode.code) {
      case 1:
        return this.appMPGS(afcFile, partnerFile, dateRange);
      case 2:
        return this.appSacombank(afcFile, partnerFile, dateRange);
      case 3:
        return this.appMomo(afcFile, partnerFile, dateRange);
      case 4:
        return this.appZalopay(afcFile, partnerFile, dateRange);
      case 5:
        return this.appShopeePay(afcFile, partnerFile, dateRange);
      case 6:
        return this.tapAndGoZaloPay(afcFile, partnerFile, dateRange);
      case 7:
        return this.tapAndGoMomo(afcFile, partnerFile, dateRange);
      case 8:
        return this.tapAndGoShopeePay(afcFile, partnerFile, dateRange);
      case 9:
        return this.kioskTTT(afcFile, partnerFile, dateRange);
      case 10:
        return this.kioskZaloPay(afcFile, partnerFile, dateRange);
      case 11:
        return this.kioskShopeePay(afcFile, partnerFile, dateRange);
      case 12:
        return this.kioskMomo(afcFile, partnerFile, dateRange);
      case 13:
        return this.kioskTechcombank(afcFile, partnerFile, dateRange);
      default:
        throw new BadRequestException(
          `Mã code ${selectedCode.code} chưa được hỗ trợ.`,
        );
    }
  }
  private filterByDateRange(
    data: DataSet,
    dateColumnIndex: number,
    dateColumn: string,
    dateRange: DateRange,
    fileLabel: string,
  ): DataSet {
    const rows = data.rows.filter((row) => {
      const value = row.values[dateColumnIndex];
      if (value === null || value === undefined || value === '') {
        throw new BadRequestException(
          `Thiếu ngày ở cột "${dateColumn}", dòng ${row.rowNumber} của ${fileLabel}.`,
        );
      }

      const rowDate = this.toCalendarDate(value);
      if (!rowDate) {
        throw new BadRequestException(
          `Ngày không hợp lệ ở cột "${dateColumn}", dòng ${row.rowNumber} của ${fileLabel}.`,
        );
      }

      return rowDate >= dateRange.fromDate && rowDate <= dateRange.toDate;
    });

    return { ...data, rows };
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

  private indexByKey(
    mainData: DataRow[],
    secondary: DataRow[],
  ): { rowsByKey: Map<string, number> } {
    const secondaryCounts = new Map<string, number>();
    for (const row of secondary) {
      secondaryCounts.set(row.key, (secondaryCounts.get(row.key) ?? 0) + 1);
    }

    const rowsByKey = new Map<string, number>();
    for (const row of mainData) {
      if (!rowsByKey.has(row.key)) {
        rowsByKey.set(row.key, secondaryCounts.get(row.key) ?? 0);
      }
    }

    return { rowsByKey };
  }
//========================================APP========================================
  private async appMomo(
    afcFile: UploadedExcelFile,
    partnerFile: UploadedExcelFile,
    dateRange: DateRange
  ) {
    const partnerSelectedCol = ['Mã đơn hàng gốc', 'Số tiền', 'Thời gian']
    const partnerKey = 'Mã đơn hàng gốc';
    const partnerPrice = 'Số tiền';
    const partner = await this.readRows(partnerFile, 'partnerFile', partnerSelectedCol, partnerKey, partnerPrice);

    const afcSelectedCol = ['Mã đối soát', 'Giá bán', 'Ngày mua vé']
    const afcKey = 'Mã đối soát';
    const afcPrice = 'Giá bán';
    const afc = await this.readRows(afcFile, 'afcFile', afcSelectedCol, afcKey, afcPrice);

    return this.recociliation(
      this.filterByDateRange(partner, 2, 'Thời gian', dateRange, 'partnerFile'),
      this.filterByDateRange(afc, 2, 'Ngày mua vé', dateRange, 'afcFile'),
    );
  }
  private async appMPGS(
    afcFile: UploadedExcelFile,
    partnerFile: UploadedExcelFile,
    dateRange: DateRange
  ) {
    const partnerSelectedCol = ['Reference_No', 'TRANS_AMOUNT', 'TRANS_DATE']
    const partnerKey = 'Reference_No';
    const partnerPrice = 'TRANS_AMOUNT';
    const partner = await this.readRows(partnerFile, 'partnerFile', partnerSelectedCol, partnerKey, partnerPrice);

    const afcSelectedCol = ['Mã đối soát', 'Giá bán', 'Ngày mua vé']
    const afcKey = 'Mã đối soát';
    const afcPrice = 'Giá bán';
    const afc = await this.readRows(afcFile, 'afcFile', afcSelectedCol, afcKey, afcPrice);

    return this.recociliation(
      this.filterByDateRange(partner, 2, 'Thời gian', dateRange, 'partnerFile'),
      this.filterByDateRange(afc, 2, 'Ngày mua vé', dateRange, 'afcFile'),
    );
  }
  private async appSacombank(
    afcFile: UploadedExcelFile,
    partnerFile: UploadedExcelFile,
    dateRange: DateRange
  ) {
    const partnerSelectedCol = ['Mã đơn hàng gốc', 'Số tiền', 'Thời gian']
    const partnerKey = 'Mã đơn hàng gốc';
    const partnerPrice = 'Số tiền';
    const partner = await this.readRows(partnerFile, 'partnerFile', partnerSelectedCol, partnerKey, partnerPrice);

    const afcSelectedCol = ['Mã đối soát', 'Giá bán', 'Ngày mua vé']
    const afcKey = 'Mã đối soát';
    const afcPrice = 'Giá bán';
    const afc = await this.readRows(afcFile, 'afcFile', afcSelectedCol, afcKey, afcPrice);

    return this.recociliation(
      this.filterByDateRange(partner, 2, 'Thời gian', dateRange, 'partnerFile'),
      this.filterByDateRange(afc, 2, 'Ngày mua vé', dateRange, 'afcFile'),
      undefined,
      undefined,
      undefined,
      true,
    );
  }
  private async appZalopay(
    afcFile: UploadedExcelFile,
    partnerFile: UploadedExcelFile,
    dateRange: DateRange
  ) {
    const partnerSelectedCol = ['Payment Request ID', 'Số tiền thanh toán trước phí', 'Ngày giao dịch']
    const partnerKey = 'Payment Request ID';
    const partnerPrice = 'Số tiền thanh toán trước phí';
    const partner = await this.readRows(partnerFile, 'partnerFile', partnerSelectedCol, partnerKey, partnerPrice);

    const afcSelectedCol = ['Mã đối soát', 'Giá bán', 'Ngày mua vé']
    const afcKey = 'Mã đối soát';
    const afcPrice = 'Giá bán';
    const afc = await this.readRows(afcFile, 'afcFile', afcSelectedCol, afcKey, afcPrice);

    return this.recociliation(
      this.filterByDateRange(partner, 2, 'Thời gian', dateRange, 'partnerFile'),
      this.filterByDateRange(afc, 2, 'Ngày mua vé', dateRange, 'afcFile'),
      undefined,
      undefined,
      true,
      true,
    );
  }
  private async appShopeePay(
    afcFile: UploadedExcelFile,
    partnerFile: UploadedExcelFile,
    dateRange: DateRange
  ) {
    const partnerSelectedCol = ['Mã đơn hàng gốc', 'Số tiền', 'Thời gian']
    const partnerKey = 'Mã đơn hàng gốc';
    const partnerPrice = 'Số tiền';
    const partner = await this.readRows(partnerFile, 'partnerFile', partnerSelectedCol, partnerKey, partnerPrice);

    const afcSelectedCol = ['Mã đối soát', 'Giá bán', 'Ngày mua vé']
    const afcKey = 'Mã đối soát';
    const afcPrice = 'Giá bán';
    const afc = await this.readRows(afcFile, 'afcFile', afcSelectedCol, afcKey, afcPrice);

    return this.recociliation(
      this.filterByDateRange(partner, 2, 'Thời gian', dateRange, 'partnerFile'),
      this.filterByDateRange(afc, 2, 'Ngày mua vé', dateRange, 'afcFile'),
    );
  }
//========================================TAP&GO========================================
  private async tapAndGoZaloPay(
    afcFile: UploadedExcelFile,
    partnerFile: UploadedExcelFile,
    dateRange: DateRange
  ) {
    const partnerSelectedCol = ['Mã đơn hàng gốc', 'Số tiền', 'Thời gian']
    const partnerKey = 'Mã đơn hàng gốc';
    const partnerPrice = 'Số tiền';
    const partner = await this.readRows(partnerFile, 'partnerFile', partnerSelectedCol, partnerKey, partnerPrice);

    const afcSelectedCol = ['Mã đối soát', 'Giá bán', 'Ngày mua vé']
    const afcKey = 'Mã đối soát';
    const afcPrice = 'Giá bán';
    const afc = await this.readRows(afcFile, 'afcFile', afcSelectedCol, afcKey, afcPrice);

    return this.recociliation(
      this.filterByDateRange(partner, 2, 'Thời gian', dateRange, 'partnerFile'),
      this.filterByDateRange(afc, 2, 'Ngày mua vé', dateRange, 'afcFile'),
    );
  }
  private async tapAndGoMomo(
    afcFile: UploadedExcelFile,
    partnerFile: UploadedExcelFile,
    dateRange: DateRange
  ) {
    const partnerSelectedCol = ['Mã đơn hàng gốc', 'Số tiền', 'Thời gian']
    const partnerKey = 'Mã đơn hàng gốc';
    const partnerPrice = 'Số tiền';
    const partner = await this.readRows(partnerFile, 'partnerFile', partnerSelectedCol, partnerKey, partnerPrice);

    const afcSelectedCol = ['Mã đối soát', 'Giá bán', 'Ngày mua vé']
    const afcKey = 'Mã đối soát';
    const afcPrice = 'Giá bán';
    const afc = await this.readRows(afcFile, 'afcFile', afcSelectedCol, afcKey, afcPrice);

    return this.recociliation(
      this.filterByDateRange(partner, 2, 'Thời gian', dateRange, 'partnerFile'),
      this.filterByDateRange(afc, 2, 'Ngày mua vé', dateRange, 'afcFile'),
    );
  }
  private async tapAndGoShopeePay(
    afcFile: UploadedExcelFile,
    partnerFile: UploadedExcelFile,
    dateRange: DateRange
  ) {
    const partnerSelectedCol = ['Mã đơn hàng gốc', 'Số tiền', 'Thời gian']
    const partnerKey = 'Mã đơn hàng gốc';
    const partnerPrice = 'Số tiền';
    const partner = await this.readRows(partnerFile, 'partnerFile', partnerSelectedCol, partnerKey, partnerPrice);

    const afcSelectedCol = ['Mã đối soát', 'Giá bán', 'Ngày mua vé']
    const afcKey = 'Mã đối soát';
    const afcPrice = 'Giá bán';
    const afc = await this.readRows(afcFile, 'afcFile', afcSelectedCol, afcKey, afcPrice);

    return this.recociliation(
      this.filterByDateRange(partner, 2, 'Thời gian', dateRange, 'partnerFile'),
      this.filterByDateRange(afc, 2, 'Ngày mua vé', dateRange, 'afcFile'),
    );
  }

//========================================KIOSK========================================
  private async kioskTTT(
    afcFile: UploadedExcelFile,
    partnerFile: UploadedExcelFile,
    dateRange: DateRange,
  ) {
    const partnerSelectedCol = ['Reference_No', 'TRANS_AMOUNT', 'TRANS_DATE']
    const partnerKey = 'Reference_No';
    const partnerPrice = 'TRANS_AMOUNT';
    const partner = await this.readRows(partnerFile, 'partnerFile', partnerSelectedCol, partnerKey, partnerPrice);

    const afcSelectedCol = ['paymentNo', 'totalAmount', 'date_created']
    const afcKey = 'paymentNo';
    const afcPrice = 'totalAmount';
    const afc = await this.readRows(afcFile, 'afcFile', afcSelectedCol, afcKey, afcPrice);

    return this.recociliation(
      this.filterByDateRange(partner, 2, 'TRANS_DATE', dateRange, 'partnerFile'),
      this.filterByDateRange(afc, 2, 'date_created', dateRange, 'afcFile'),
    );
  }
  
  private async kioskShopeePay(afcFile: UploadedExcelFile, partnerFile: UploadedExcelFile, dateRange: DateRange) {
    const partnerSelectedCol = ['Mã tham chiếu', 'Giá trị giao dịch', 'Thời gian tạo']
    const partnerKey = 'Mã tham chiếu';
    const partnerPrice = 'Giá trị giao dịch';
    const partner = await this.readRows(partnerFile, 'partnerFile', partnerSelectedCol, partnerKey, partnerPrice);

    const afcSelectedCol = ['id', 'totalAmount', 'date_created']
    const afcKey = 'id';
    const afcPrice = 'totalAmount';
    const afc = await this.readRows(afcFile, 'afcFile', afcSelectedCol, afcKey, afcPrice);

    return this.recociliation(
      this.filterByDateRange(partner, 2, 'Thời gian tạo', dateRange, 'partnerFile'),
      this.filterByDateRange(afc, 2, 'date_created', dateRange, 'afcFile'),
      'O',
    );
  }

  private async kioskZaloPay(afcFile: UploadedExcelFile, partnerFile: UploadedExcelFile, dateRange: DateRange) {
    const partnerSelectedCol = ['Mã GD Zalopay', 'Số tiền thanh toán trước phí', 'Ngày giao dịch']
    const partnerKey = 'Mã GD Zalopay';
    const partnerPrice = 'Số tiền thanh toán trước phí';
    const partner = await this.readRows(partnerFile, 'partnerFile', partnerSelectedCol, partnerKey, partnerPrice);

    const afcSelectedCol = ['id', 'totalAmount', 'date_created']
    const afcKey = 'id';
    const afcPrice = 'totalAmount';
    const afc = await this.readRows(afcFile, 'afcFile', afcSelectedCol, afcKey, afcPrice);
    const partnerWithoutFailedTransactions = {
      ...partner,
      rows: partner.rows.filter((row) => {
        const status = row.values[3];
        return !(
          typeof status === 'string' &&
          status.trim().toLocaleLowerCase('vi-VN').includes('thất bại')
        );
      }),
    };

    return this.recociliation(
      this.filterByDateRange(
        partnerWithoutFailedTransactions,
        2,
        'Ngày giao dịch',
        dateRange,
        'partnerFile',
      ),
      this.filterByDateRange(afc, 2, 'date_created', dateRange, 'afcFile'),
    );
  }

  private async kioskMomo(afcFile: UploadedExcelFile, partnerFile: UploadedExcelFile, dateRange: DateRange) {
    const partnerSelectedCol = ['Bill_ID', 'Total_Amount', 'Date']
    const partnerKey = 'Bill_ID';
    const partnerPrice = 'Total_Amount';
    const partner = await this.readRows(partnerFile, 'partnerFile', partnerSelectedCol, partnerKey, partnerPrice);

    const afcSelectedCol = ['id', 'total_amount', 'datecreated']
    const afcKey = 'id';
    const afcPrice = 'total_amount';
    const afc = await this.readRows(afcFile, 'afcFile', afcSelectedCol, afcKey, afcPrice);

    return this.recociliation(
      this.filterByDateRange(partner, 2, 'Date', dateRange, 'partnerFile'),
      this.filterByDateRange(afc, 2, 'datecreated', dateRange, 'afcFile'),
      undefined,
      'O',
    );
  }

  private async kioskTechcombank(afcFile: UploadedExcelFile, partnerFile: UploadedExcelFile, dateRange: DateRange) {
    const partnerSelectedCol = ['Mã đơn hàng', 'Số tiền đơn hàng', 'Thời gian giao dịch']
    const partnerKey = 'Mã đơn hàng';
    const partnerPrice = 'Số tiền đơn hàng';
    const partner = await this.readRows(partnerFile, 'partnerFile', partnerSelectedCol, partnerKey, partnerPrice);

    const afcSelectedCol = ['id', 'totalAmount', 'date_created']
    const afcKey = 'id';
    const afcPrice = 'totalAmount';
    const afc = await this.readRows(afcFile, 'afcFile', afcSelectedCol, afcKey, afcPrice);

    return this.recociliation(
      this.filterByDateRange(partner, 2, 'Thời gian giao dịch', dateRange, 'partnerFile'),
      this.filterByDateRange(afc, 2, 'date_created', dateRange, 'afcFile'),
      'O',
    );
  }

  private async recociliation(
    partnerData: DataSet,
    afcData: DataSet,
    addKeyAFC?: string,
    addKeyPartner?: string,
    cutKeyAFC = false,
    cutKeyPartner = false,
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

    return Buffer.from(await result.xlsx.writeBuffer());
  }

  private findUnmatchedRows(
    sourceRows: DataRow[],
    counterpartRows: DataRow[],
    sourceAddKey?: string,
    counterpartAddKey?: string,
    sourceCutKeyAtSpace = false,
    counterpartCutKeyAtSpace = false,
  ): {
    unmatchedRows: DataRow[];
    extraRows: DataRow[];
  } {
    const addPrefix = (key: string, prefix?: string): string => {
      const normalizedPrefix = prefix?.toLocaleLowerCase('vi-VN');
      return normalizedPrefix && !key.startsWith(normalizedPrefix)
        ? `${normalizedPrefix}${key}`
        : key;
    };
    const normalizeKey = (key: string, cutAtSpace: boolean): string =>
      cutAtSpace ? key.trim().split(/\s+/, 1)[0] : key;

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
    transactionTypeColumn?: string,
  ): Promise<DataSet> {
    if (!file.originalname.toLowerCase().endsWith('.xlsx')) {
      throw new BadRequestException(`${fileLabel} phải là file Excel .xlsx.`);
    }

    const rows: DataRow[] = [];
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
        let transactionTypeColumnIndex: number | undefined;
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
            transactionTypeColumnIndex = transactionTypeColumn
              ? findColumn(transactionTypeColumn)
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
          rows.push(
            this.toDataRow(
              row,
              keyColumnIndex,
              priceColumnIndex,
              selectedColumnIndexes,
              transactionTypeColumnIndex,
              fullHeaderCount,
            ),
          );
        }

        return {
          headers: [...selectedColumns],
          fullHeaders,
          rows,
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
    transactionTypeColumn?: number,
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
      transactionType: transactionTypeColumn
        ? row.getCell(transactionTypeColumn).text
        : undefined,
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
