import { Module } from '@nestjs/common';
import { DbModule } from '../db/db.module';
import { ReconciliationController } from './reconciliation.controller';
import { CronJobService } from './cron-job.service';
import { ReconciliationService } from './reconciliation.service';
import { B2StorageService } from './B2Storage.service';

@Module({
  imports: [DbModule],
  controllers: [ReconciliationController],
  providers: [ReconciliationService,B2StorageService],//CronJobService
})
export class ReconciliationModule {}
