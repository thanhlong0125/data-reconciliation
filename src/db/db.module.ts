import { Module } from '@nestjs/common';
import { MongoHttpExceptionFilter } from './mongo-http-exception.filter';
import { MongoLogService } from './mongo-log.service';

@Module({
  providers: [MongoLogService, MongoHttpExceptionFilter],
  exports: [MongoLogService, MongoHttpExceptionFilter],
})
export class DbModule {}
