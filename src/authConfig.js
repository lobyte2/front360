// ---------------------------------------------------------------------------
// AZURE ENTRA ID (INSTITUCIONAL)
// ---------------------------------------------------------------------------
const getRedirectUri = () => {
    return window.location.hostname === "localhost"
        ? window.location.origin
        : "https://jfgd1pckb9.execute-api.us-east-1.amazonaws.com";
};

export const msalConfig = {
    auth: {
        clientId: "4a72e40f-1d5a-497a-b040-d2fa5acb61fd", // ID de aplicación Frontend
        authority: "https://login.microsoftonline.com/dd140621-3398-4bfe-805b-9c162df51706", // ID del Tenant
        redirectUri: getRedirectUri(),
        postLogoutRedirectUri: getRedirectUri()
    },
    cache: {
        cacheLocation: "sessionStorage", 
        storeAuthStateInCookie: false, 
    }
};

export const loginRequest = {
    scopes: [
        "api://cb8b1711-d599-4a45-94d9-9251f15db61a/write", 
        "api://cb8b1711-d599-4a45-94d9-9251f15db61a/read"
    ]
};

// ---------------------------------------------------------------------------
// AWS COGNITO (CLIENTES / EXTERNOS)
// ---------------------------------------------------------------------------
export const cognitoAuthConfig = {
    authority: "https://cognito-idp.us-east-1.amazonaws.com/us-east-1_EPwV5Yl4n",
    client_id: "5qipn3rrr5j8uajmmsthq9i0cs",
    redirect_uri: getRedirectUri(),
    response_type: "code",
    scope: "phone openid email",
};

// Dominio Hosted UI de Cognito
export const cognitoDomain = "https://us-east-1epwv5yl4n.auth.us-east-1.amazoncognito.com";

export const cognitoLogoutUri = getRedirectUri();

export const signOutRedirect = () => {
    const clientId = "5qipn3rrr5j8uajmmsthq9i0cs";
    const logoutUri = getRedirectUri();
    window.location.href = `${cognitoDomain}/logout?client_id=${clientId}&logout_uri=${encodeURIComponent(logoutUri)}`;
};

export const cognitoOidcConfig = {
    ...cognitoAuthConfig,
    post_logout_redirect_uri: getRedirectUri(),
    loadUserInfo: false,
    monitorSession: false,
    automaticSilentRenew: false
};

// ---------------------------------------------------------------------------
// Sincronización entre Azure y Cognito
// ---------------------------------------------------------------------------
export const PENDING_AUTH_PROVIDER_KEY = "pending_auth_provider";
export const AUTH_PROVIDER_AZURE = "azure";
export const AUTH_PROVIDER_COGNITO = "cognito";