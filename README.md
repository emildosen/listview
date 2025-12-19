<p align="center">
<img src="public/banner.png" width="300">
</p>
<br>

# ListView

ListView is a web app that lets you manage data stored in SharePoint lists. Like Notion, you build what you need without being locked into a specific use case: tracking, dashboards, workflows, and more.

<br>

## Features

<table>
<tr>
<td align="center" width="25%"><b>Schema-driven UI</b></td>
<td align="center" width="25%"><b>Works with any list</b></td>
<td align="center" width="25%"><b>Pure SPA</b></td>
<td align="center" width="25%"><b>M365 styling</b></td>
</tr>
<tr>
<td>Components discover list structure via Graph metadata, no hardcoding required</td>
<td>Point it at an existing SharePoint list and start working immediately</td>
<td>No backend, all Graph calls from the browser using delegated permissions</td>
<td>Fluent UI components for a consistent Microsoft experience</td>
</tr>
</table>

<br>

## Getting Started

**Go to [listview.org](https://listview.org)** and sign in with your Microsoft 365 account.

Works with any commercial M365 tenant. Your tenant admin may need to grant consent for the app before you can sign in.

<br>

## Self-Hosting

Self-host ListView if you need:

- Custom Entra ID app registration (your own client ID, custom permissions)
- Support for GCC, GCC High, or other sovereign clouds

### Prerequisites

- Node.js 18+
- Entra ID app registration with delegated permissions:
  - Microsoft Graph: `Sites.ReadWrite.All`, `User.Read`
  - SharePoint: `AllSites.Manage`

### Setup

1. **Register an Entra ID app**

   - Go to Entra ID admin center > App registrations > New registration
   - Set redirect URI to `http://localhost` (SPA platform)

2. **Clone and configure**

   ```bash
   git clone https://github.com/emildosen/listview.git
   cd listview
   cp .env.example .env
   ```

   Edit `.env`:

   ```dotenv
   VITE_MSAL_CLIENT_ID=your-client-id-here
   ```

3. **Run**

   ```bash
   npm install
   npm run dev
   ```

<br>

## Tech Stack

- React + TypeScript
- Vite
- MSAL for authentication
- Microsoft Graph and PnP for SharePoint operations
- Fluent UI

<br>

## Architecture

- **No backend** - All authentication and data access happens client-side
- **Config-driven** - The app adapts to whatever lists you point it at
- **Tenant-wide settings** - Stored in a dedicated `sites/ListView` SharePoint site

<br>

## Contributing

Contributions welcome.

<br>

## License

[MIT](LICENSE.md)
