export const msalConfig = {
    auth: {
        clientId: "4a72e40f-1d5a-497a-b040-d2fa5acb61fd", // ID de aplicación Frontend
        authority: "https://login.microsoftonline.com/dd140621-3398-4bfe-805b-9c162df51706", // ID del Tenant
        redirectUri: "http://localhost:5173",
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