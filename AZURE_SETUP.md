# Azure App Registration Setup for MCP Outlook Scheduler

## Overview
This guide will help you set up Azure App Registration with the correct permissions to access Outlook calendars through Microsoft Graph API.

## ⚠️ IMPORTANT: Permission Types

### Application Permissions (App-only)
- **What it is**: The app acts on its own behalf, not on behalf of a user
- **Use case**: Background services, daemons, admin operations
- **Limitation**: Cannot access user-specific data like calendars

### Delegated Permissions (User-based)
- **What it is**: The app acts on behalf of a signed-in user
- **Use case**: User calendar access, meeting scheduling, availability checking
- **Required for**: This MCP server to work properly

## 🔧 Step-by-Step Setup

### 1. Create Azure App Registration
1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **New registration**
4. Fill in:
   - **Name**: `MCP Outlook Scheduler`
   - **Supported account types**: `Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)`
   - **Redirect URI**: `http://localhost:3000/auth` (Web platform)

### 2. Configure API Permissions
1. In your app registration, go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Choose **Delegated permissions**
5. Add these permissions:
   - **Calendars.ReadWrite** - Read and write user calendars
   - **User.Read** - Read user profile
   - **MailboxSettings.Read** - Read user working hours
   - **Calendars.ReadWrite.Shared** - Access shared calendars
   - **OnlineMeetings.ReadWrite** - Create Teams meetings

### 3. Grant Admin Consent
1. Click **Grant admin consent for [Your Organization]**
2. Confirm the action
3. Verify all permissions show **Granted for [Your Organization]**

### 4. Configure Authentication for Device Code Flow
1. Go to **Authentication** in your app registration
2. Under **Platform configurations**, click **Add a platform**
3. Select **Mobile and desktop applications**
4. Check **Enable the following mobile and desktop flows**:
   - ✅ **Device code flow**
5. Click **Configure**
6. **IMPORTANT**: Add this redirect URI: `https://login.microsoftonline.com/common/oauth2/nativeclient`
7. **CRITICAL**: Enable **"Allow public client flows"** in the Authentication settings
8. **SAVE CHANGES**: Make sure to click the **Save** button to persist your authentication configuration

### 5. Get Application Credentials
1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Add description: `MCP Server Secret`
4. **IMPORTANT**: Copy the **Value** (not the ID) immediately
5. Go to **Overview** and copy:
   - **Application (client) ID**
   - **Directory (tenant) ID**

### 6. Update Environment Variables
Update your `.env` file:

```env
# Change this from 'app' to 'delegated'
GRAPH_AUTH_MODE=delegated

# Your Azure app registration details
GRAPH_CLIENT_ID=your-client-id-here
GRAPH_TENANT_ID=your-tenant-id-here
GRAPH_CLIENT_SECRET=your-client-secret-value-here

# Your email (must be in the same tenant)
GRAPH_ORGANIZER_EMAIL=your-email@yourdomain.com
```

## 🔍 Why This Fixes the Issue

### Current Problem
- `GRAPH_AUTH_MODE=app` tries to use application permissions
- Application permissions cannot access user calendars
- Results in "Failed to create event" and "Failed to get schedule" errors

### Solution
- `GRAPH_AUTH_MODE=delegated` uses delegated permissions
- Delegated permissions can access user calendars
- Requires user consent (device code flow) but works for calendar operations

## 🚀 Testing

After updating your `.env` file:

1. **Rebuild the project**:
   ```bash
   npm run build
   ```

2. **Test authentication**:
   ```bash
   node test-auth-only.js
   ```

3. **Test availability**:
   ```bash
   node test-availability.js
   ```

4. **Test booking**:
   ```bash
   node test-direct-booking.js
   ```

## 📝 Notes

- **First run**: You'll need to authenticate via device code (one-time setup)
- **Subsequent runs**: Token will be cached in `.tokens.json`
- **Token expiration**: Tokens expire and need renewal (handled automatically)
- **User consent**: Only needed once per user per app

## 🆘 Troubleshooting

### Still getting "Failed to create event"?
1. Verify `GRAPH_AUTH_MODE=delegated` in `.env`
2. Check that admin consent was granted
3. Ensure your email is in the same Azure tenant
4. Verify all required permissions are granted

### Permission denied errors?
1. Check admin consent status
2. Verify permission names match exactly
3. Ensure you're using the correct tenant ID
4. Check if your account has admin rights in the tenant

### "invalid_client" error in delegated mode?
1. **Add redirect URI**: `https://login.microsoftonline.com/common/oauth2/nativeclient`
2. **Enable device code flow** in Authentication settings
3. **Supported account types**: Should include personal Microsoft accounts
4. **Clear token cache**: Delete `.tokens.json` file
