import 'reflect-metadata';
import { NestFactory } from '@nestjs/core';
import { ExpressAdapter } from '@nestjs/platform-express';
import type { Express } from 'express';
import type { IncomingMessage, ServerResponse } from 'node:http';
import { AppModule } from '../src/app.module';

let server: Express | undefined;

async function createServer(): Promise<Express> {
  if (server) {
    return server;
  }

  const expressApp: Express = require('express')();
  const app = await NestFactory.create(
    AppModule,
    new ExpressAdapter(expressApp),
  );
  app.enableCors();
  await app.init();

  server = expressApp;
  return expressApp;
}

export default async function handler(
  request: IncomingMessage,
  response: ServerResponse,
) {
  const expressApp = await createServer();
  if (request.url?.startsWith('/api')) {
    request.url = '/reconciliation/readfile';
  }
  return expressApp(request, response);
}
