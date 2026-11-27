import type { Document, ObjectId } from 'mongodb';

export interface UploadedFileDocument extends Document {
  _id?: ObjectId;
  name: string;
  links: string[];
  status: string;
  key: string;
  userKey: string;
  size: number;
  contentType: string;
  uploadedAt: Date;
}

export interface NewUploadedFileDocument {
  name: string;
  links: string[];
  status: string;
  key: string;
  userKey: string | undefined;
  size: number;
  contentType: string;
}
