# Gtoolbox Context & Guidelines

## Runtime & Plataforma
- Google Apps Script (V8 runtime).
- Sync tool: `clasp`.
- APIs de Node.js (`fs`, `path`, `process`) y módulos ES6 (`import`/`export`) no soportados directamente por el runtime.
- Usar servicios estándar de Apps Script (`SpreadsheetApp`, `DriveApp`, `GmailApp`, `UrlFetchApp`, `Logger`, etc.).

## Ciclo de Desarrollo
1. Modificaciones locales en `.js` / `.gs`.
2. Subida a la nube: `clasp push`.
3. Sincronización a GitHub: commits claros y descriptivos.
