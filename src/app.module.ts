import { Module } from '@nestjs/common';
import { ConfigModule } from '@nestjs/config';
import { ReconciliationModule } from './reconciliation/reconciliation.module';

@Module({
  imports: [
    ConfigModule.forRoot({ isGlobal: true }),
    ReconciliationModule,
  ],
})
export class AppModule {}
