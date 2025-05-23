import * as dotenv from 'dotenv';

dotenv.config();

export const config = {
  tenantId: process.env.TENANT_ID || '',
  clientId: process.env.CLIENT_ID || '',
  clientSecret: process.env.CLIENT_SECRET || '',
  redirectUri: process.env.REDIRECT_URI || 'http://localhost:3000/auth/callback',
  port: parseInt(process.env.PORT || '3000'),
  testUserEmail: process.env.TEST_USER_EMAIL || '',
};

// Validate required config
const requiredFields = ['tenantId', 'clientId', 'clientSecret'];
for (const field of requiredFields) {
  if (!config[field as keyof typeof config]) {
    throw new Error(`Missing required environment variable: ${field.toUpperCase()}`);
  }
}