import type { Document, ObjectId } from 'mongodb';

export interface SystemLogDocument extends Document {
  _id?: ObjectId;
  event: string;
  details: Document;
  createdAt: Date;
}
