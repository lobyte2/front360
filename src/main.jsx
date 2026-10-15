import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App.jsx';
import './index.css';

import { PublicClientApplication, EventType } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import { AuthProvider } from "react-oidc-context";
import {
  msalConfig,
  cognitoOidcConfig,
  PENDING_AUTH_PROVIDER_KEY,
  AUTH_PROVIDER_COGNITO,
} from "./authConfig";

const msalInstance = new PublicClientApplication(msalConfig);

msalInstance.addEventCallback((event) => {
  if (
    event.eventType === EventType.LOGIN_FAILURE ||
    event.eventType === EventType.ACQUIRE_TOKEN_FAILURE
  ) {
    console.error("Fallo de MSAL:", event.error);
  }
});

function hayCallbackDeAuth(search = window.location.search) {
  const params = new URLSearchParams(search);
  return Boolean((params.get("code") || params.get("error")) && params.get("state"));
}

function limpiarUrl() {
  const url = new URL(window.location.href);
  ["code", "state", "session_state", "error", "error_description", "error_uri"]
    .forEach((param) => url.searchParams.delete(param));
  window.history.replaceState(null, "", url.pathname + url.search + url.hash);
}

async function bootstrap() {
  await msalInstance.initialize();

  const proveedorPendiente = sessionStorage.getItem(PENDING_AUTH_PROVIDER_KEY);
  const callbackDeCognito = proveedorPendiente === AUTH_PROVIDER_COGNITO;
  const callbackDeAzure = hayCallbackDeAuth() && !callbackDeCognito;

  if (callbackDeAzure) {
    try {
      const response = await msalInstance.handleRedirectPromise();
      if (response) {
        console.info("Autenticacion Azure completada:", response.account?.username);
      }
    } catch (error) {
      console.error("Error al procesar la respuesta de Azure:", error);
    }
    limpiarUrl();
  } else if (!callbackDeCognito) {
    limpiarUrl();
  }

  sessionStorage.removeItem(PENDING_AUTH_PROVIDER_KEY);

  ReactDOM.createRoot(document.getElementById('root')).render(
    <React.StrictMode>
      <AuthProvider
        {...cognitoOidcConfig}
        skipSigninCallback={!callbackDeCognito}
        onSigninCallback={limpiarUrl}
      >
        <MsalProvider instance={msalInstance}>
          <App />
        </MsalProvider>
      </AuthProvider>
    </React.StrictMode>
  );
}

bootstrap().catch((error) => {
  console.error("Error al inicializar autenticación:", error);
});