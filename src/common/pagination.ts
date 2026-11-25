import { BadRequestException } from '@nestjs/common';

const DEFAULT_PAGE = 1;
const DEFAULT_LIMIT = 10;
const MAX_LIMIT = 100;

export interface PaginationOptions {
  page: number;
  limit: number;
  skip: number;
}

export interface PaginatedResult<T> {
  data: T[];
  page: number;
  limit: number;
  total: number;
  totalPages: number;
}

export function parsePagination(
  pageValue?: string,
  limitValue?: string,
): PaginationOptions {
  const page = parsePositiveInteger(pageValue, DEFAULT_PAGE, 'page');
  const limit = parsePositiveInteger(limitValue, DEFAULT_LIMIT, 'limit');

  if (limit > MAX_LIMIT) {
    throw new BadRequestException(`limit không được lớn hơn ${MAX_LIMIT}.`);
  }

  const skip = (page - 1) * limit;
  if (!Number.isSafeInteger(skip)) {
    throw new BadRequestException('page vượt quá giới hạn cho phép.');
  }

  return { page, limit, skip };
}

export function createPaginatedResult<T>(
  data: T[],
  total: number,
  pagination: PaginationOptions,
): PaginatedResult<T> {
  return {
    data,
    page: pagination.page,
    limit: pagination.limit,
    total,
    totalPages: Math.ceil(total / pagination.limit),
  };
}

function parsePositiveInteger(
  value: string | undefined,
  defaultValue: number,
  fieldName: string,
): number {
  if (value === undefined) {
    return defaultValue;
  }

  const parsedValue = Number(value);
  if (
    !/^\d+$/.test(value) ||
    !Number.isSafeInteger(parsedValue) ||
    parsedValue < 1
  ) {
    throw new BadRequestException(`${fieldName} phải là số nguyên dương.`);
  }

  return parsedValue;
}
