import {
  Injectable,
  Logger,
  OnModuleDestroy,
  OnModuleInit,
} from '@nestjs/common';
import { MongoClient } from 'mongodb';
import type { Db, Document } from 'mongodb';

type LogDetails = Document;

const LOG_RETENTION_SECONDS = 10 * 24 * 60 * 60;

@Injectable()
export class MongoLogService implements OnModuleInit, OnModuleDestroy {
  private readonly logger = new Logger(MongoLogService.name);
  private client: MongoClient | undefined;
  private database: Db | undefined;

  async onModuleInit(): Promise<void> {
    console.log("process.env.MONGODB_URI ", process.env.MONGODB_URI)
    const uri = process.env.MONGODB_URI?.trim();
    console.log("uri", uri)
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
      await this.recordSystemEvent('application_started');
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
