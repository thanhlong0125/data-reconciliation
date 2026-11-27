import type { Document, ObjectId } from 'mongodb';

export interface UserLogDocument extends Document {
  _id?: ObjectId;
  userKey: string;
  action: string;
  details: Document;
  createdAt: Date;
}
