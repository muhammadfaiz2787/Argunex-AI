import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "./index.css";
import App from "./App.jsx";

import { PublicClientApplication } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";

const msalConfig = {
  auth: {
    clientId: "3cbfd24a-638a-4603-a19d-80256b34a51e",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://argunex-ai-ten.vercel.app",
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
  system: {
    allowRedirectInIframe: false,
  }
};

const msalInstance = new PublicClientApplication(msalConfig);

msalInstance.initialize().then(() => {
  // TANGANI REDIRECT DARI LOGIN (kunci agar tidak perlu popup)
  msalInstance.handleRedirectPromise()
    .then(() => {
      createRoot(document.getElementById("root")).render(
        <StrictMode>
          <MsalProvider instance={msalInstance}>
            <App />
          </MsalProvider>
        </StrictMode>
      );
    })
    .catch((error) => {
      console.error("Gagal menangani redirect MSAL:", error);
      // Tetap render app meski ada error redirect
      createRoot(document.getElementById("root")).render(
        <StrictMode>
          <MsalProvider instance={msalInstance}>
            <App />
          </MsalProvider>
        </StrictMode>
      );
    });
}).catch((error) => {
  console.error("Gagal menginisialisasi MSAL:", error);
});