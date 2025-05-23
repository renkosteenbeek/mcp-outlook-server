export interface AccountConfig {
  name: string;
  tenantId: string;
  clientId: string;
  clientSecret: string;
}

export interface ServerConfig {
  port: number;
  redirectUri: string;
}

export interface Config {
  accounts: AccountConfig[];
  server: ServerConfig;
}

export interface AccountTokenCache {
  [accountName: string]: any;
}

export interface AccountResponse<T> {
  account: string;
  data: T;
  error?: string;
}