import type { Document, ObjectId } from 'mongodb';

export interface CutKeyConfig {
  key: string;
  index: number;
}

export interface ReconciliationConfigDocument extends Document {
  _id?: ObjectId;
  code: number;
  name: string;
  partner: {
    key: string;
    price: string;
    time: string;
    refund?: string;
    addKey?: string;
    cutKeyPartner?: CutKeyConfig;
  };
  afc: {
    key: string;
    price: string;
    time: string;
    addKey?: string;
    cutKeyPartner?: CutKeyConfig;
  };
  ignoreFailedPartnerTransactions?: boolean;
  active: boolean;
}
