import tl = require('azure-pipelines-task-lib');
import { Client } from '@microsoft/microsoft-graph-client';
import * as fs from 'fs';
import * as path from 'path';
import 'isomorphic-fetch';

// Required for Microsoft Graph Client
require('isomorphic-fetch');

export function isNullOrEmpty(str: string | null | undefined): boolean {
    return str === null || str === undefined || str.trim() === '';
}

export async function getAuthToken(clientId: string, clientSecret: string, tenantId: string): Promise<string> {
    const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    
    const data = new URLSearchParams();
    data.append('client_id', clientId);
    data.append('scope', 'https://graph.microsoft.com/.default');
    data.append('client_secret', clientSecret);
    data.append('grant_type', 'client_credentials');
    
    try {
        const response = await fetch(tokenEndpoint, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            },
            body: data
        });
        
        const responseData = await response.json();
        
        if (!response.ok) {
            throw new Error(`Authentication failed: ${responseData.error_description || responseData.error || 'Unknown error'}`);
        }
        
        return responseData.access_token;
    } catch (error) {
        console.error('Error getting auth token:', error);
        throw error;
    }
}

export async function createGraphClient(accessToken: string): Promise<Client> {
    return Client.init({
        authProvider: (done) => {
            done(null, accessToken);
        }
    });
}

export async function run() {
    try {
        const To: string | undefined = tl.getInputRequired('To');
        const Subject: string | undefined = tl.getInputRequired('Subject');
        const From: string | undefined = tl.getInputRequired('From');
        const AppId: string | undefined = tl.getInputRequired('AppId');
        const ClientSecret: string | undefined = tl.getInputRequired('ClientSecret');
        const TenantId: string | undefined = tl.getInputRequired('TenantId');
        
        const Body: string | undefined = tl.getInput('Body');
        const BodyAsHtml: boolean | undefined = tl.getBoolInput('BodyAsHtml');
        const AddAttachment: boolean | undefined = tl.getBoolInput('AddAttachment');
        const Attachment: string | undefined = tl.getInput('Attachment');
        const CC: string | undefined = tl.getInput('CC');
        const BCC: string | undefined = tl.getInput('BCC');
        
        console.log('Input Values');
        console.log('To:', To);
        console.log('Subject:', Subject);
        console.log('Body:', Body);
        console.log('From:', From);
        console.log('AppId:', AppId);
        console.log('TenantId:', TenantId);
        // Not logging ClientSecret for security reasons
        console.log('AddAttachment:', AddAttachment);
        console.log('Attachment:', Attachment);
        console.log('CC:', CC);
        console.log('BCC:', BCC);

        // Check required fields
        if (!To || !Subject || !From || !AppId || !ClientSecret || !TenantId) {
            console.log('Required fields are missing');
            tl.setResult(tl.TaskResult.Failed, 'Required fields are missing');
            return;
        }

        // Get auth token
        const accessToken = await getAuthToken(AppId, ClientSecret, TenantId);
        console.log('Successfully obtained access token');

        // Create Graph client
        const client = await createGraphClient(accessToken);
        console.log('Graph client initialized');

        // Prepare email message
        const message: any = {
            message: {
                subject: Subject,
                body: {
                    contentType: BodyAsHtml ? 'HTML' : 'Text',
                    content: Body || ''
                },
                toRecipients: To.split(',').map(email => ({
                    emailAddress: {
                        address: email.trim()
                    }
                }))
            },
            saveToSentItems: true
        };

        // Add CC recipients if provided
        if (!isNullOrEmpty(CC)) {
            message.message.ccRecipients = CC.split(',').map(email => ({
                emailAddress: {
                    address: email.trim()
                }
            }));
        }

        // Add BCC recipients if provided
        if (!isNullOrEmpty(BCC)) {
            message.message.bccRecipients = BCC.split(',').map(email => ({
                emailAddress: {
                    address: email.trim()
                }
            }));
        }

        // Add attachment if specified
        if (AddAttachment && !isNullOrEmpty(Attachment)) {
            try {
                const fileName = path.basename(Attachment);
                const fileContent = fs.readFileSync(Attachment);
                const base64Content = fileContent.toString('base64');

                message.message.attachments = [
                    {
                        '@odata.type': '#microsoft.graph.fileAttachment',
                        name: fileName,
                        contentBytes: base64Content
                    }
                ];
                
                console.log(`Attachment ${fileName} added`);
            } catch (err) {
                console.error(`Error adding attachment: ${err}`);
                tl.setResult(tl.TaskResult.Failed, `Failed to add attachment: ${err}`);
                return;
            }
        }

        // Send email using Microsoft Graph API
        console.log('Sending email...');
        await client.api(`/users/${From}/sendMail`).post(message);
        
        console.log('Email sent successfully');
        tl.setResult(tl.TaskResult.Succeeded, 'Email sent successfully');
    } catch (err) {
        console.error('Error:', err);
        tl.setResult(tl.TaskResult.Failed, `Failed to send email: ${err}`);
    }
}

run();