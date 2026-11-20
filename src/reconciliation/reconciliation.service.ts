import { BadRequestException, Injectable } from '@nestjs/common';
import ExcelJS from 'exceljs';
import { Readable } from 'node:stream';
import { UploadedExcelFile } from './uploaded-excel-file.type';

const PARTNER_KEY = 'Mã đơn hàng gốc';
const PARTNER_PRICE = 'Số tiền';
const AFC_KEY = 'Mã đối soát';
const AFC_PRICE = 'Giá bán';
const AFC_SELECTED_COLUMNS = [
  'Mã đối soát',
  'Ngày mua vé',
  'Giá bán',
  'Mã vé',
];
const PARTNER_SELECTED_COLUMNS = [
  'Mã đơn hàng gốc',
  'Thời gian',
  'Loại giao dịch',
  'Số tiền',
];
const CODE = [
  // { code: 1, name: 'App - Thẻ quốc thế (MPGS)'},
  // { code: 2, name: 'App - TKTT Sacombank' },
  // { code: 3, name: 'App - Momo' },
  // { code: 4, name: 'App - Zalopay' },
  // { code: 5, name: 'App - ShopeePay' },
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
  getCodes(): { code: number; name: string }[] {
    return CODE.map(({ code, name }) => ({ code, name }));
  }

  async checkCodeAndHandle(
    afcFile: UploadedExcelFile,
    partnerFile: UploadedExcelFile,
    code: number,
    dateRange?: DateRange,
  ) {
    const selectedCode = CODE.find(item => item.code === code);
    if (!selectedCode) {
      throw new BadRequestException(`Mã code ${code} không hợp lệ.`);
    }
    if (!dateRange) {
      throw new BadRequestException('Vui lòng chọn khoảng ngày đối soát.');
    }
    switch (selectedCode.code) {
      case 1:
      case 2:
      case 3:
        return this.appMomo(afcFile, partnerFile, dateRange);
      case 4:
      case 5:
      case 6:
      case 7:
      case 8:
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
        : value.toISOString().slice(0, 10);
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
          Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000,
        );
        return Number.isNaN(excelDate.getTime())
          ? undefined
          : excelDate.toISOString().slice(0, 10);
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
    if (!timezone) {
      return dateText;
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
    const timestamp = new Date(
      `${dateText}T${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}:${String(second).padStart(2, '0')}.${(millisecondText ?? '0').padEnd(3, '0')}${normalizedTimezone}`,
    );
    return Number.isNaN(timestamp.getTime())
      ? undefined
      : timestamp.toISOString().slice(0, 10);
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
    );
  }
  private async appZalopay(
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
  private async ShopeePay(
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
  ): Promise<Buffer> {

    const partnerRowsByMatch = this.findUnmatchedRows(
      partnerData.rows,
      afcData.rows,
      addKeyPartner,
      addKeyAFC,
    );
    const afcRowsByMatch = this.findUnmatchedRows(
      afcData.rows,
      partnerData.rows,
      addKeyAFC,
      addKeyPartner,
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

  async checkNotMatchingID(
    afcFile: UploadedExcelFile,
    partnerFile: UploadedExcelFile,
    code?: string,
  ): Promise<Buffer> {
    void code;
    const afc = await this.readRows(
      afcFile,
      'afcFile',
      AFC_SELECTED_COLUMNS,
      AFC_KEY,
      AFC_PRICE,
    );

    const result = new ExcelJS.Workbook();
    const partnerData = await this.Refunds(result, partnerFile);
    const partnerHeaders = partnerData.remaining.headers;
    const partnerRefunds = partnerData.refunds;

    const partnerRowsByKey = this.indexByKey(
      partnerData.remaining.rows,
      afc.rows,
    );
    const afcRowsByKey = this.indexByKey(
      afc.rows,
      partnerData.remaining.rows,
    );
    const wrongPriceRowsByKey = this.wrongPriceRowsByKey(
      afc.rows,
      partnerData.remaining.rows,
    );

    const partnerRowsByMatch = this.findUnmatchedRows(
      partnerData.remaining.rows,
      afc.rows,
    );
    const afcRowsByMatch = this.findUnmatchedRows(
      afc.rows,
      partnerData.remaining.rows,
    );
    const partnerUnmatched = [
      ...partnerRowsByMatch.unmatchedRows,
      ...partnerRowsByMatch.extraRows,
    ];
    const afcUnmatched = [
      ...afcRowsByMatch.unmatchedRows,
      ...afcRowsByMatch.extraRows,
    ];
    const partnerUnmatchedSheet = result.addWorksheet('Partner dư');
    partnerUnmatchedSheet.addRow(partnerData.remaining.fullHeaders);
    for (const row of partnerUnmatched) {
      partnerUnmatchedSheet.addRow(row.fullValues);
    }

    const afcUnmatchedSheet = result.addWorksheet('AFC dư');
    afcUnmatchedSheet.addRow(afc.fullHeaders);
    for (const row of afcUnmatched) {
      afcUnmatchedSheet.addRow(row.fullValues);
    }

    const summarySheet = result.addWorksheet('Tong hop');
    this.addSummaryRows(summarySheet, partnerUnmatched, afcUnmatched);

    const wrongPriceSheet = result.addWorksheet('Lech gia');
    wrongPriceSheet.addRow([
      'Mã đối soát',
      ...afc.headers.map((header) => `AFC - ${header}`),
      ...partnerRefunds.headers.map((header) => `Partner - ${header}`),
      'Giá lệch (AFC - Partner)',
    ]);
    for (const { afcRow, partnerRow, priceDifference } of wrongPriceRowsByKey.values()) {
      wrongPriceSheet.addRow([
        afcRow.displayKey,
        ...afcRow.values,
        ...partnerRow.values,
        priceDifference,
      ]);
    }

    const resultSheet = result.addWorksheet('Doi soat 1-1');
    resultSheet.addRow([
      'Số dòng khớp ở file đối ứng',
      'Kết quả 1-1',
      'Nguồn dữ liệu',
      'Mã đối soát',
      ...afc.headers.map((header) => `AFC - ${header}`),
      ...partnerHeaders.map((header) => `Partner - ${header}`),
    ]);

    const afcPadding = Array.from(
      { length: partnerHeaders.length },
      () => null,
    );
    const partnerPadding = Array.from({ length: afc.headers.length }, () => null);

    for (const row of afc.rows) {
      const matchStatus = this.getMatchStatus(
        row.key,
        afcRowsByKey.rowsByKey,
        partnerRowsByKey.rowsByKey,
        'afcFile',
      );
      resultSheet.addRow([
        'AFC',
        matchStatus.check,
        matchStatus.message,
        row.displayKey,
        ...row.values,
        ...afcPadding,
      ]);
    }

    for (const row of partnerRefunds.rows) {
      const matchStatus = this.getMatchStatus(
        row.key,
        afcRowsByKey.rowsByKey,
        partnerRowsByKey.rowsByKey,
        'partnerFile',
      );
      resultSheet.addRow([
        'Partner - Hoàn tiền',
        matchStatus.check,
        matchStatus.message,
        row.displayKey,
        ...partnerPadding,
        ...row.values,
      ]);
    }

    const output = Buffer.from(await result.xlsx.writeBuffer());

    return output;
  }

  private findUnmatchedRows(
    sourceRows: DataRow[],
    counterpartRows: DataRow[],
    sourceAddKey?: string,
    counterpartAddKey?: string,
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
    const counterpartCountsByKey = new Map<string, number>();
    for (const row of counterpartRows) {
      if (!row.key) {
        continue;
      }
      const key = addPrefix(row.key, counterpartAddKey);
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

      const key = addPrefix(row.key, sourceAddKey);
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
  
  private addSummaryRows(
    summarySheet: ExcelJS.Worksheet,
    partnerUnmatched: DataRow[],
    afcUnmatched: DataRow[],
  ): void {
    const partnerTotal = partnerUnmatched.reduce(
      (total, row) => total + row.price,
      0,
    );
    const afcTotal = afcUnmatched.reduce(
      (total, row) => total + row.price,
      0,
    );

    summarySheet.addRow([
      'STT',
      'Giải trình',
      'Chi tiết giao dịch tại',
    ]);
    summarySheet.addRow([
      1,
      `AFC có ${afcUnmatched.length} giao dịch trị giá ${afcTotal} VND mà Partner không có (thu bù)`,
      'AFC dư',
    ]);
    summarySheet.addRow([
      2,
      `Partner có ${partnerUnmatched.length} giao dịch trị giá ${partnerTotal} VND mà AFC không có (hoàn tiền)`,
      'Partner dư',
    ]);
  }

  private async Refunds(
    workbook: ExcelJS.Workbook,
    partnerFile: UploadedExcelFile,
  ): Promise<RefundData> {
    // Partner key and transaction type are part of PARTNER_SELECTED_COLUMNS.
    const partnerData = await this.readRows(
      partnerFile,
      'partnerFile',
      PARTNER_SELECTED_COLUMNS,
      PARTNER_KEY,
      PARTNER_PRICE,
      'Loại giao dịch',
    );
    const partnerRefundRows: DataRow[] = [];
    const partnerRemainingRows: DataRow[] = [];

    for (const row of partnerData.rows) {
      const isRefund = (row.transactionType ?? '')
        .toLocaleLowerCase('vi-VN')
        .includes('hoàn tiền');
      (isRefund ? partnerRefundRows : partnerRemainingRows).push(row);
    }

    const partnerRefunds: DataSet = {
      headers: partnerData.headers,
      fullHeaders: partnerData.fullHeaders,
      rows: partnerRefundRows,
    };
    const partnerRemaining: DataSet = {
      headers: partnerData.headers,
      fullHeaders: partnerData.fullHeaders,
      rows: partnerRemainingRows,
    };

    const refundSheet = workbook.addWorksheet('Hoan tien');
    refundSheet.addRow(partnerRefunds.headers);
    for (const row of partnerRefunds.rows) {
      refundSheet.addRow(row.values);
    }

    return {
      refunds: partnerRefunds,
      remaining: partnerRemaining,
    };
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

  private wrongPriceRowsByKey(
    afcRows: DataRow[],
    partnerRows: DataRow[],
  ): Map<string, WrongPriceRow> {
    const rowsByKey = new Map<string, WrongPriceRow>();
    const firstPartnerRowByKey = new Map<string, DataRow>();
    const firstDifferentPriceRowByKey = new Map<string, DataRow>();

    for (const partnerRow of partnerRows) {
      const firstPartnerRow = firstPartnerRowByKey.get(partnerRow.key);
      if (!firstPartnerRow) {
        firstPartnerRowByKey.set(partnerRow.key, partnerRow);
      } else if (
        !firstDifferentPriceRowByKey.has(partnerRow.key) &&
        partnerRow.price !== firstPartnerRow.price
      ) {
        firstDifferentPriceRowByKey.set(partnerRow.key, partnerRow);
      }
    }

    for (const afcRow of afcRows) {
      if (rowsByKey.has(afcRow.key)) {
        continue;
      }

      const firstPartnerRow = firstPartnerRowByKey.get(afcRow.key);
      const partnerRow =
        firstPartnerRow?.price !== afcRow.price
          ? firstPartnerRow
          : firstDifferentPriceRowByKey.get(afcRow.key);
      if (partnerRow) {
        rowsByKey.set(afcRow.key, {
          afcRow,
          partnerRow,
          priceDifference: afcRow.price - partnerRow.price,
        });
      }
    }

    return rowsByKey;
  }

  private getMatchStatus(
    key: string,
    afcRowsByKey: Map<string, number>,
    partnerRowsByKey: Map<string, number>,
    sourceFile: 'afcFile' | 'partnerFile',
  ): { check: number; message: string } {
    if (!key) {
      return { check: 0, message: 'Không đạt - Thiếu Mã đối soát' };
    }

    const afcCount = afcRowsByKey.get(key) ?? 0;
    const partnerCount = partnerRowsByKey.get(key) ?? 0;
    const counterpartCount =
      sourceFile === 'afcFile' ? partnerCount : afcCount;
    if (afcCount === 1 && partnerCount === 1) {
      return { check: counterpartCount, message: 'Khớp lệnh' };
    }

    const issues: string[] = [];
    if (afcCount === 0) {
      issues.push('không tìm thấy ở afcFile');
    } else if (afcCount > 1) {
      issues.push(`trùng ${afcCount} dòng ở afcFile`);
    }

    if (partnerCount === 0) {
      issues.push('không tìm thấy ở partnerFile');
    } else if (partnerCount > 1) {
      issues.push(`trùng ${partnerCount} dòng ở partnerFile`);
    }
    if (afcCount > partnerCount) {
      issues.push(`dư ${afcCount - partnerCount} dòng ở afcFile`);
    } else if (partnerCount > afcCount) {
      issues.push(`dư ${partnerCount - afcCount} dòng ở partnerFile`);
    }

    return {
      check: afcCount,
      message: `Không đạt - ${issues.join('; ')} (AFC ${afcCount}, partner ${partnerCount})`,
    };
  }

}
