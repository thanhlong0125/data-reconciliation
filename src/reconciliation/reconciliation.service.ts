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
  {code: 1, name:'App - Thẻ quốc thế (MPGS)', partnerKey: 'Mã đơn hàng gốc', afcKey: 'Mã đối soát'},
  {code: 2, name:'App - TKTT Sacombank', partnerKey: 'Mã đơn hàng gốc', afcKey: 'Mã đối soát'},
  {code: 3, name:'App - Momo', partnerKey: 'Mã đơn hàng gốc', afcKey: 'Mã đối soát'},
  {code: 4, name:'App - Zalopay', partnerKey: 'Mã đơn hàng gốc', afcKey: 'Mã đối soát'},
  {code: 5, name:'Tap&go - ZaloPay', partnerKey: 'Mã đơn hàng gốc', afcKey: 'Mã đối soát'},
  {code: 6, name:'Tap&go - Momo', partnerKey: 'Mã đơn hàng gốc', afcKey: 'Mã đối soát'},
  {code: 7, name:'Kiosk - TTT', partnerKey: 'Mã đơn hàng gốc', afcKey: 'Mã đối soát'},
  {code: 8, name:'Kiosk - ZaloPay', partnerKey: 'Mã đơn hàng gốc', afcKey: 'Mã đối soát'},
  {code: 9, name:'App - ShopeePay', partnerKey: 'Mã đơn hàng gốc', afcKey: 'Mã đối soát'},
  {code: 10, name:'Tap&go - ShopeePay', partnerKey: 'Mã đơn hàng gốc', afcKey: 'Mã đối soát'},
  {code: 11, name:'Kiosk - ShopeePay', partnerKey: 'Mã đơn hàng gốc', afcKey: 'Mã đối soát'},
]


interface DataRow {
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

interface CodeData {
  code: number;
  name: string;
  partnerKey: string;
  afcKey: string;
}

@Injectable()
export class ReconciliationService {

  async checkCodeAndHandle(
    afcFile: UploadedExcelFile,
    partnerFile: UploadedExcelFile,
    code: number) {
      const selectedCode = CODE.find(item => item.code == code);
      if (!selectedCode) {
        throw new BadRequestException(`Mã code ${code} không hợp lệ.`);
      }
      switch (selectedCode.code) {
        case 1:
        case 2:
        case 3:
        case 4:
          return this.checkNotMatchingID(afcFile, partnerFile, selectedCode.name, );
        case 5:
        case 6:
          return this.checkNotMatchingID(afcFile, partnerFile, selectedCode.name);
        case 7:
          return this.kioskTTT(afcFile, partnerFile, selectedCode);
        case 8:
        case 9:
        case 10:
        case 11:
          return this.checkNotMatchingID(afcFile, partnerFile, selectedCode.name);
        } 
  }


  private async kioskTTT(afcFile: UploadedExcelFile, partnerFile: UploadedExcelFile, codedata:CodeData) {
    const partnerSelectedCol = ['Reference_No','TRANS_AMOUNT','TRANS_DATE']
    const partnerKey = 'Reference_No';
    const partnerPrice = 'TRANS_AMOUNT';
    const partner = await this.readRows(partnerFile, 'partnerFile', partnerSelectedCol, partnerKey, partnerPrice);


    const afcSelectedCol = ['paymentNo','totalAmount','date_created']
    const afcKey = 'paymentNo';
    const afcPrice = 'totalAmount';
    const afc = await this.readRows(afcFile, 'afcFile', afcSelectedCol, afcKey, afcPrice);

    return this.recociliation(partner, afc);
  }

  private async recociliation(
    partnerData: DataSet,
    afcData: DataSet,
  ): Promise<Buffer> {
    const partnerUnmatched = this.findUnmatchedRows(
      partnerData.rows,
      afcData.rows,
    );
    const afcUnmatched = this.findUnmatchedRows(
      afcData.rows,
      partnerData.rows,
    );

    const result = new ExcelJS.Workbook();

    const partnerUnmatchedSheet = result.addWorksheet('Partner dư');
    partnerUnmatchedSheet.addRow(partnerData.fullHeaders);
    for (const row of partnerUnmatched) {
      partnerUnmatchedSheet.addRow(row.fullValues);
    }

    const afcUnmatchedSheet = result.addWorksheet('AFC dư');
    afcUnmatchedSheet.addRow(afcData.fullHeaders);
    for (const row of afcUnmatched) {
      afcUnmatchedSheet.addRow(row.fullValues);
    }

    const summarySheet = result.addWorksheet('Tong hop');
    this.addSummaryRows(summarySheet, partnerUnmatched, afcUnmatched);

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

    const partnerUnmatched = this.findUnmatchedRows(
      partnerData.remaining.rows,
      afc.rows,
    );
    const afcUnmatched = this.findUnmatchedRows(
      afc.rows,
      partnerData.remaining.rows,
    );
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
  ): DataRow[] {
    const counterpartCountsByKey = new Map<string, number>();
    for (const row of counterpartRows) {
      if (!row.key) {
        continue;
      }
      counterpartCountsByKey.set(
        row.key,
        (counterpartCountsByKey.get(row.key) ?? 0) + 1,
      );
    }

    const sourceOccurrencesByKey = new Map<string, number>();
    const unmatchedRows: DataRow[] = [];
    for (const row of sourceRows) {
      if (!row.key) {
        unmatchedRows.push(row);
        continue;
      }

      const sourceOccurrence =
        (sourceOccurrencesByKey.get(row.key) ?? 0) + 1;
      sourceOccurrencesByKey.set(row.key, sourceOccurrence);

      const counterpartCount = counterpartCountsByKey.get(row.key) ?? 0;
      if (sourceOccurrence > counterpartCount) {
        unmatchedRows.push(row);
      }
    }

    return unmatchedRows;
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
    const displayKey = row.getCell(keyColumn).text.trim();
    return {
      key: displayKey.toLocaleLowerCase('vi-VN'),
      displayKey,
      price: parseFloat(row.getCell(priceColumn).text.trim()) || 0,
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

  private countByKey(rows: DataRow[]): Map<string, number> {
    const counts = new Map<string, number>();
    for (const row of rows) {
      if (row.key) {
        counts.set(row.key, (counts.get(row.key) ?? 0) + 1);
      }
    }
    return counts;
  }

  private wrongPriceRowsByKey(
    afcRows: DataRow[],
    partnerRows: DataRow[],
  ): Map<string, WrongPriceRow> {
    const rowsByKey = new Map<string, WrongPriceRow>();
    // Cũ: partnerRows.find(...) chạy lại từ đầu cho từng dòng AFC, cũng có thể thành O(n x m).
    // const partnerRow = partnerRows.find(
    //   (row) => row.key === afcRow.key && row.price !== afcRow.price,
    // );
    // Mới: lập chỉ mục theo key trong O(n), mỗi lần tra cứu dòng partner là O(1).
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
