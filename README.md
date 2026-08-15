# OneDrive Excel Dashboard

Dashboard Next.js que lee una planilla de Excel en tu OneDrive personal (cuenta Microsoft Live) vía Microsoft Graph, con login OAuth.

## Stack

- Next.js (App Router) + TypeScript + Tailwind
- Auth.js / NextAuth con Microsoft Entra ID (`consumers` = cuentas personales)
- Microsoft Graph: workbook / worksheets / usedRange
- Deploy: Vercel

## Excel cableada

Item ID (resid de OneDrive):

`3cec66d5-ed44-4799-8089-a801fc93a9f2`

Configurable con `EXCEL_DRIVE_ITEM_ID`.

## 1. App Registration en Azure (obligatorio)

1. Entrá a [Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. **New registration**
   - Name: `onedrive-excel-dashboard`
   - Supported account types: **Personal Microsoft accounts only** (o “any org + personal”)
   - Redirect URI (Web):
     - Local: `http://localhost:3000/api/auth/callback/microsoft-entra-id`
     - Prod: `https://<tu-app>.vercel.app/api/auth/callback/microsoft-entra-id`
3. Certificates & secrets → **New client secret** → copiá el Value
4. API permissions → Add → Microsoft Graph (Delegated):
   - `User.Read`
   - `Files.Read`
   - `openid`, `profile`, `email`, `offline_access` (suelen ir implícitos en OIDC)
5. Copiá **Application (client) ID**

## 2. Variables de entorno

Copiá `.env.example` a `.env.local` y completá:

```bash
AUTH_SECRET=...          # openssl rand -base64 32
AUTH_URL=http://localhost:3000
AUTH_MICROSOFT_ENTRA_ID_ID=...
AUTH_MICROSOFT_ENTRA_ID_SECRET=...
AUTH_MICROSOFT_ENTRA_ID_ISSUER=https://login.microsoftonline.com/consumers/v2.0
EXCEL_DRIVE_ITEM_ID=3cec66d5-ed44-4799-8089-a801fc93a9f2
```

En Vercel usá las mismas keys (con `AUTH_URL` = URL de producción).

## 3. Desarrollo local

```bash
npm install
npm run dev
```

Abrí http://localhost:3000 → **Iniciar sesión con Microsoft / OneDrive**.

## 4. Deploy Vercel

```bash
npx vercel --prod
```

O conectá el repo en el dashboard de Vercel y cargá las env vars.
