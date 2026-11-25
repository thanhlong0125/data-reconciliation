
import { Injectable, Logger } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import {
    S3Client,
    GetObjectCommand,
    DeleteObjectCommand,
} from '@aws-sdk/client-s3';
import { Upload } from '@aws-sdk/lib-storage';
import { getSignedUrl } from '@aws-sdk/s3-request-presigner';
import { randomUUID } from 'crypto';
import { Readable } from 'node:stream';

interface MemoryUploadedFile {
    originalname: string;
    mimetype: string;
    size: number;
    buffer: Uint8Array;
}

@Injectable()
export class B2StorageService {
    private readonly logger = new Logger(B2StorageService.name);
    private readonly s3: S3Client;
    private readonly bucket: string;

    constructor(private readonly config: ConfigService) {
        this.bucket = this.config.getOrThrow<string>('B2_BUCKET_NAME');
        const endpoint = this.config.getOrThrow<string>('B2_ENDPOINT');
        const endpointRegion = new URL(endpoint).hostname.match(
            /^s3\.([^.]+)\.backblazeb2\.com$/,
        )?.[1];
        const region = this.config.get<string>('B2_REGION') ?? endpointRegion;
        if (!region) {
            throw new Error(
                'B2_REGION is required when it cannot be derived from B2_ENDPOINT.',
            );
        }

        this.s3 = new S3Client({
            endpoint,
            region,
            credentials: {
                accessKeyId: this.config.getOrThrow<string>('B2_KEY_ID'),
                secretAccessKey: this.config.getOrThrow<string>(
                    'B2_APPLICATION_KEY',
                ),
            },
        });
    }

    async uploadFile(file: MemoryUploadedFile) {
        try {

            const key = `temp/${randomUUID()}-${file.originalname}`;

            const upload = new Upload({
                client: this.s3,
                params: {
                    Bucket: this.bucket,
                    Key: key,
                    Body: file.buffer,
                    ContentType: file.mimetype,
                },
                partSize: 10 * 1024 * 1024,
                queueSize: 2,
            });

            await upload.done();

            const endpoint = this.config
                .getOrThrow<string>('B2_ENDPOINT')
                .replace(/\/+$/, '');
            const links = `${endpoint}/${encodeURIComponent(this.bucket)}/${key
                .split('/')
                .map(encodeURIComponent)
                .join('/')}`;

            return {
                key,
                url: links,
                originalName: file.originalname,
                size: file.size,
                contentType: file.mimetype,
            };
        } catch (error) {
            const message =
                error instanceof Error ? error.message : String(error);
            this.logger.error(
                `B2 upload failed: ${message}`,
                error instanceof Error ? error.stack : undefined,
            );
            throw new Error(`Lỗi khi upload file lên B2: ${message}`);
        }
    }
    async deleteFile(key: string) {
        await this.s3.send(
            new DeleteObjectCommand({
                Bucket: this.bucket,
                Key: key,
            }),
        );

        return { success: true };
    }

    async downloadFile(key: string): Promise<Readable> {
        const result = await this.s3.send(
            new GetObjectCommand({
                Bucket: this.bucket,
                Key: key,
            }),
        );

        if (!result.Body) {
            throw new Error('File not found');
        }

        return result.Body as Readable;
}
    async getDownloadUrl(key: string) {
        const command = new GetObjectCommand({
            Bucket: this.bucket,
            Key: key,
        });

        return getSignedUrl(this.s3, command, {
            expiresIn: 900,
        });
    }
}
