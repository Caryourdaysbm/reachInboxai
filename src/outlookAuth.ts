import * as msal from '@azure/msal-node';

const msalConfig = {
  auth: {
    clientId: process.env.OUTLOOK_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.OUTLOOK_TENANT_ID}`,
    clientSecret: process.env.OUTLOOK_CLIENT_SECRET,
  },
};

const pca = new msal.ConfidentialClientApplication(msalConfig);

export const getOutlookAuthUrl = async () => {
  const authCodeUrlParameters = {
    scopes: ['https://graph.microsoft.com/.default'],
    redirectUri: process.env.OUTLOOK_REDIRECT_URL,
  };

  return await pca.getAuthCodeUrl(authCodeUrlParameters);
};

export const getOutlookTokens = async (code: string) => {
  const tokenResponse = await pca.acquireTokenByCode({
    code,
    scopes: ['https://graph.microsoft.com/.default'],
    redirectUri: process.env.OUTLOOK_REDIRECT_URL,
  });

  return tokenResponse;
};
