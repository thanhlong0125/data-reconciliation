import {
  BadRequestException,
  InternalServerErrorException,
  Injectable,
  Logger,
  NotFoundException,
  OnModuleDestroy,
  OnModuleInit,
} from '@nestjs/common';
import { MongoClient, ObjectId } from 'mongodb';
import type { Db, Document } from 'mongodb';
import { createPaginatedResult } from '../common/pagination';
import type {
  PaginatedResult,
  PaginationOptions,
} from '../common/pagination';

type LogDetails = Document;

const LOG_RETENTION_SECONDS = 10 * 24 * 60 * 60;

@Injectable()
export class MongoLogService implements OnModuleInit, OnModuleDestroy {
  private readonly logger = new Logger(MongoLogService.name);
  private client: MongoClient | undefined;
  private database: Db | undefined;

  async onModuleInit(): Promise<void> {
    const uri = process.env.MONGODB_URI?.trim();
    if (!uri) {
      this.logger.warn('MONGODB_URI chưa cấu hình; log MongoDB đang tắt.');
      return;
    }

    const client = new MongoClient(uri, { serverSelectionTimeoutMS: 10_000 });
    this.client = client;
    try {
      await client.connect();
      const database = client.db(
        process.env.MONGODB_DATABASE?.trim() || 'data-reconciliation',
      );
      this.database = database;
      await Promise.all([
        database
          .collection('user_logs')
          .createIndex(
            { createdAt: 1 },
            { expireAfterSeconds: LOG_RETENTION_SECONDS },
          ),
        database
          .collection('system_logs')
          .createIndex(
            { createdAt: 1 },
            { expireAfterSeconds: LOG_RETENTION_SECONDS },
          ),
      ]);
      this.logger.log('Đã kết nối MongoDB và bật lưu log.');
    } catch (error) {
      await client.close();
      this.client = undefined;
      this.database = undefined;
      throw error;
    }
  }

  async onModuleDestroy(): Promise<void> {
    const client = this.client;
    if (!client) {
      return;
    }

    await this.recordSystemEvent('application_stopped');
    await client.close();
    this.client = undefined;
    this.database = undefined;
  }
  async getUploadedFiles(
    userKey: string | undefined,
    pagination: PaginationOptions,
  ): Promise<PaginatedResult<Document>> {
    if (!this.database) {
      throw new InternalServerErrorException(
        'MongoDB chưa kết nối; không thể lấy danh sách file đã upload.',
      );
    }

    const collection = this.database.collection('uploaded_files');
    const filter = { userKey: userKey || 'unknown' };
    const [files, total] = await Promise.all([
      collection
        .find(filter)
        .sort({ uploadedAt: -1, _id: -1 })
        .skip(pagination.skip)
        .limit(pagination.limit)
        .toArray(),
      collection.countDocuments(filter),
    ]);

    return createPaginatedResult(files, total, pagination);
  }

  async recordUserAction(
    userKey: string | undefined,
    action: string,
    details: LogDetails,
  ): Promise<void> {
    await this.insertLog('user_logs', {
      userKey: userKey || 'unknown',
      action,
      details,
      createdAt: new Date(),
    });
  }

  async recordSystemEvent(
    event: string,
    details: LogDetails = {},
  ): Promise<void> {
    await this.insertLog('system_logs', {
      event,
      details,
      createdAt: new Date(),
    });
  }

  async recordUploadedFile(record: {
    name: string;
    links: string[];
    status: string;
    key: string;
    userKey: string | undefined;
    size: number;
    contentType: string;
  }): Promise<string> {
    if (!this.database) {
      throw new Error(
        'MongoDB chưa kết nối; không thể lưu thông tin file đã upload.',
      );
    }

    const result = await this.database.collection('uploaded_files').insertOne({
      ...record,
      userKey: record.userKey || 'unknown',
      uploadedAt: new Date(),
    });
    return result.insertedId.toString();
  }

  async getUploadedFile(id: string): Promise<{
    key: string;
    originalName: string;
    mimeType: string;
  }> {
    if (!ObjectId.isValid(id)) {
      throw new BadRequestException('ID file không hợp lệ.');
    }
    if (!this.database) {
      throw new InternalServerErrorException(
        'MongoDB chưa kết nối; không thể tìm file.',
      );
    }

    const file = await this.database
      .collection('uploaded_files')
      .findOne({ _id: new ObjectId(id) });
    if (!file) {
      throw new NotFoundException('Không tìm thấy file đã upload.');
    }
    if (
      typeof file.key !== 'string' ||
      typeof file.name !== 'string' ||
      typeof file.contentType !== 'string'
    ) {
      throw new InternalServerErrorException(
        'Thông tin file lưu trong MongoDB không hợp lệ.',
      );
    }

    return {
      key: file.key,
      originalName: file.name,
      mimeType: file.contentType,
    };
  }

  private async insertLog(
    collectionName: 'user_logs' | 'system_logs',
    log: LogDetails,
  ): Promise<void> {
    if (!this.database) {
      return;
    }

    try {
      await this.database.collection(collectionName).insertOne(log);
    } catch (error) {
      const message =
        error instanceof Error ? error.message : 'Lỗi không xác định.';
      this.logger.error(`Không ghi được ${collectionName}: ${message}`);
    }
  }
}
