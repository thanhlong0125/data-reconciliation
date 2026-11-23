import {
  ArgumentsHost,
  Catch,
  ExceptionFilter,
  HttpException,
  Injectable,
} from '@nestjs/common';
import type { Request, Response } from 'express';
import { MongoLogService } from './mongo-log.service';

@Catch()
@Injectable()
export class MongoHttpExceptionFilter implements ExceptionFilter {
  constructor(private readonly mongoLogService: MongoLogService) {}

  async catch(exception: unknown, host: ArgumentsHost): Promise<void> {
    const context = host.switchToHttp();
    const request = context.getRequest<Request>();
    const response = context.getResponse<Response>();
    const statusCode =
      exception instanceof HttpException ? exception.getStatus() : 500;
    const message =
      exception instanceof Error ? exception.message : 'Lỗi không xác định.';
    const isUploadValidationError =
      statusCode < 500 &&
      (request.path === '/reconciliation/upload' ||
        request.path === '/reconciliation/readfile');

    await this.mongoLogService.recordSystemEvent(
      isUploadValidationError ? 'upload_validation_error' : 'api_error',
      {
        method: request.method,
        path: request.path,
        statusCode,
        message,
        userKey: request.ip,
      },
    );

    if (exception instanceof HttpException) {
      response.status(statusCode).json(exception.getResponse());
      return;
    }

    response
      .status(statusCode)
      .json({ statusCode, message: 'Internal server error' });
  }
}
