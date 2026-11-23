import 'reflect-metadata';
import { NestFactory } from '@nestjs/core';
import type { NestExpressApplication } from '@nestjs/platform-express';
import { existsSync } from 'node:fs';
import { join } from 'node:path';
import { AppModule } from './app.module';
import { MongoHttpExceptionFilter } from './db/mongo-http-exception.filter';

async function bootstrap(): Promise<void> {
  const envFile = join(process.cwd(), '.env');
  if (existsSync(envFile)) {
    process.loadEnvFile(envFile);
  }

  const app = await NestFactory.create<NestExpressApplication>(AppModule);
  app.set('trust proxy', 1);
  app.useGlobalFilters(app.get(MongoHttpExceptionFilter));
  app.enableShutdownHooks();
  app.useStaticAssets(join(__dirname, '..', 'public'));
  await app.listen(process.env.PORT ?? 3000);
}

void bootstrap();
