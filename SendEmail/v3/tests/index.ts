import * as tl from 'azure-pipelines-task-lib/task';
import { isNullOrEmpty, getAuthToken, createGraphClient, run } from '../index';

// Mock dependencies
let mockAuthToken = 'mock-auth-token';
let mockGraphClient = {
    api: jest.fn().mockReturnThis(),
    post: jest.fn().mockResolvedValue({ value: 'success' })
};

// Mock the external modules
jest.mock('isomorphic-fetch', () => jest.fn());
jest.mock('@microsoft/microsoft-graph-client', () => ({
    Client: {
        init: jest.fn(() => mockGraphClient)
    }
}));

// Mock task library
jest.mock('azure-pipelines-task-lib/task', () => ({
    getInputRequired: jest.fn(),
    getInput: jest.fn(),
    getBoolInput: jest.fn(),
    setResult: jest.fn(),
    TaskResult: { Failed: 'Failed', Succeeded: 'Succeeded' }
}));

// Mock fs methods
jest.mock('fs', () => ({
    readFileSync: jest.fn(() => Buffer.from('mock-file-content'))
}));

// Mock fetch for auth token
global.fetch = jest.fn(() => 
    Promise.resolve({
        ok: true,
        json: () => Promise.resolve({ access_token: mockAuthToken })
    })
) as jest.Mock;

describe('SendEmail V3 Utility Functions', () => {
    describe('isNullOrEmpty', () => {
        it('should return true for null values', () => {
            expect(isNullOrEmpty(null)).toBe(true);
        });

        it('should return true for undefined values', () => {
            expect(isNullOrEmpty(undefined)).toBe(true);
        });

        it('should return true for empty strings', () => {
            expect(isNullOrEmpty('')).toBe(true);
        });

        it('should return true for whitespace-only strings', () => {
            expect(isNullOrEmpty('   ')).toBe(true);
        });

        it('should return false for non-empty strings', () => {
            expect(isNullOrEmpty('test')).toBe(false);
        });
    });

    describe('getAuthToken', () => {
        it('should retrieve an auth token successfully', async () => {
            const token = await getAuthToken('client-id', 'client-secret', 'tenant-id');
            expect(token).toBe(mockAuthToken);
            expect(global.fetch).toHaveBeenCalledWith(
                'https://login.microsoftonline.com/tenant-id/oauth2/v2.0/token',
                expect.objectContaining({
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
                })
            );
        });

        it('should throw an error when authentication fails', async () => {
            (global.fetch as jest.Mock).mockImplementationOnce(() => Promise.resolve({
                ok: false,
                json: () => Promise.resolve({ error: 'invalid_client', error_description: 'Invalid client credentials' })
            }));

            await expect(getAuthToken('invalid-id', 'invalid-secret', 'tenant-id'))
                .rejects.toThrow('Authentication failed: Invalid client credentials');
        });
    });

    describe('createGraphClient', () => {
        it('should create a graph client with the provided token', async () => {
            const client = await createGraphClient(mockAuthToken);
            expect(client).toBe(mockGraphClient);
        });
    });
});

describe('SendEmail V3 Task', () => {
    beforeEach(() => {
        // Reset all mocks before each test
        jest.clearAllMocks();
        
        // Reset mockGraphClient implementation
        mockGraphClient.api = jest.fn().mockReturnThis();
        mockGraphClient.post = jest.fn().mockResolvedValue({ value: 'success' });
        
        // Set up mock inputs
        const mockInputs: { [key: string]: string } = {
            'To': 'test@example.com',
            'Subject': 'Test Subject',
            'From': 'sender@example.com',
            'Body': 'Test Body',
            'AppId': 'mock-app-id',
            'ClientSecret': 'mock-client-secret',
            'TenantId': 'mock-tenant-id'
        };
        
        (tl.getInputRequired as jest.Mock).mockImplementation(name => mockInputs[name]);
        (tl.getInput as jest.Mock).mockImplementation(name => mockInputs[name] || '');
        (tl.getBoolInput as jest.Mock).mockImplementation(name => false);
    });

    it('should successfully send a basic email', async () => {
        await run();

        // Verify auth token was requested
        expect(global.fetch).toHaveBeenCalled();
        
        // Verify Graph client was initialized
        expect(mockGraphClient.api).toHaveBeenCalledWith('/users/sender@example.com/sendMail');
        expect(mockGraphClient.post).toHaveBeenCalledWith({
            message: {
                subject: 'Test Subject',
                body: {
                    contentType: 'Text',
                    content: 'Test Body'
                },
                toRecipients: [
                    {
                        emailAddress: {
                            address: 'test@example.com'
                        }
                    }
                ]
            },
            saveToSentItems: true
        });

        // Verify success status was set
        expect(tl.setResult).toHaveBeenCalledWith(tl.TaskResult.Succeeded, 'Email sent successfully');
    });

    it('should send an email with HTML body when BodyAsHtml is true', async () => {
        (tl.getBoolInput as jest.Mock).mockImplementation(name => name === 'BodyAsHtml');

        await run();

        expect(mockGraphClient.post).toHaveBeenCalledWith(
            expect.objectContaining({
                message: expect.objectContaining({
                    body: {
                        contentType: 'HTML',
                        content: 'Test Body'
                    }
                })
            })
        );
    });

    it('should add CC recipients when provided', async () => {
        (tl.getInput as jest.Mock).mockImplementation(name => {
            if (name === 'CC') return 'cc1@example.com,cc2@example.com';
            return 'Test Body';
        });

        await run();

        expect(mockGraphClient.post).toHaveBeenCalledWith(
            expect.objectContaining({
                message: expect.objectContaining({
                    ccRecipients: [
                        { emailAddress: { address: 'cc1@example.com' } },
                        { emailAddress: { address: 'cc2@example.com' } }
                    ]
                })
            })
        );
    });

    it('should add BCC recipients when provided', async () => {
        (tl.getInput as jest.Mock).mockImplementation(name => {
            if (name === 'BCC') return 'bcc1@example.com,bcc2@example.com';
            return 'Test Body';
        });

        await run();

        expect(mockGraphClient.post).toHaveBeenCalledWith(
            expect.objectContaining({
                message: expect.objectContaining({
                    bccRecipients: [
                        { emailAddress: { address: 'bcc1@example.com' } },
                        { emailAddress: { address: 'bcc2@example.com' } }
                    ]
                })
            })
        );
    });

    it('should add attachments when AddAttachment is true', async () => {
        (tl.getBoolInput as jest.Mock).mockImplementation(name => name === 'AddAttachment');
        (tl.getInput as jest.Mock).mockImplementation(name => {
            if (name === 'Attachment') return '/path/to/attachment.txt';
            return 'Test Body';
        });

        const fs = require('fs');
        (fs.readFileSync as jest.Mock).mockReturnValue(Buffer.from('attachment-content'));

        await run();

        expect(mockGraphClient.post).toHaveBeenCalledWith(
            expect.objectContaining({
                message: expect.objectContaining({
                    attachments: [
                        {
                            '@odata.type': '#microsoft.graph.fileAttachment',
                            name: 'attachment.txt',
                            contentBytes: 'YXR0YWNobWVudC1jb250ZW50' // Base64 encoded 'attachment-content'
                        }
                    ]
                })
            })
        );
    });

    it('should fail when required fields are missing', async () => {
        (tl.getInputRequired as jest.Mock).mockImplementationOnce(() => { throw new Error('Input required and not supplied: To'); });

        await run();

        expect(tl.setResult).toHaveBeenCalledWith(tl.TaskResult.Failed, expect.stringContaining('Failed to send email:'));
    });

    it('should handle authentication errors', async () => {
        (global.fetch as jest.Mock).mockImplementationOnce(() => Promise.reject(new Error('Authentication failed')));

        await run();

        expect(tl.setResult).toHaveBeenCalledWith(tl.TaskResult.Failed, expect.stringContaining('Failed to send email:'));
    });

    it('should handle Graph API errors', async () => {
        mockGraphClient.post = jest.fn().mockRejectedValueOnce(new Error('Graph API error'));

        await run();

        expect(tl.setResult).toHaveBeenCalledWith(tl.TaskResult.Failed, expect.stringContaining('Failed to send email:'));
    });

    it('should handle attachment errors', async () => {
        (tl.getBoolInput as jest.Mock).mockImplementation(name => name === 'AddAttachment');
        (tl.getInput as jest.Mock).mockImplementation(name => {
            if (name === 'Attachment') return '/path/to/missing-file.txt';
            return 'Test Body';
        });

        const fs = require('fs');
        (fs.readFileSync as jest.Mock).mockImplementationOnce(() => { throw new Error('File not found'); });

        await run();

        expect(tl.setResult).toHaveBeenCalledWith(tl.TaskResult.Failed, expect.stringContaining('Failed to add attachment:'));
    });

    it('should handle multiple recipients in To field', async () => {
        (tl.getInputRequired as jest.Mock).mockImplementation(name => {
            if (name === 'To') return 'test1@example.com,test2@example.com';
            return {
                'Subject': 'Test Subject',
                'From': 'sender@example.com',
                'AppId': 'mock-app-id',
                'ClientSecret': 'mock-client-secret',
                'TenantId': 'mock-tenant-id'
            }[name];
        });

        await run();

        expect(mockGraphClient.post).toHaveBeenCalledWith(
            expect.objectContaining({
                message: expect.objectContaining({
                    toRecipients: [
                        { emailAddress: { address: 'test1@example.com' } },
                        { emailAddress: { address: 'test2@example.com' } }
                    ]
                })
            })
        );
    });
});