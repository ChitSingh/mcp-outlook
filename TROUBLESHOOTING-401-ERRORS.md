# Troubleshooting Microsoft Graph API 401 Errors

## 🚨 Problem Summary
Your MCP Outlook Scheduler is experiencing 401 (Unauthorized) errors when calling Microsoft Graph API endpoints, despite successful authentication and token generation.

## 🔍 Root Cause Analysis

Based on the logs and code analysis, the issue appears to be one of these:

1. **Token Scope Issues**: The token is being generated but lacks the required permissions
2. **Admin Consent**: Permissions are granted but admin consent hasn't been given
3. **Permission Mismatch**: The requested scopes don't match the granted permissions
4. **Organization Policies**: Conditional access or other policies blocking access

## 🛠️ Diagnostic Tools

### 1. Run the Comprehensive Diagnostic
```bash
node diagnose-graph-api.js
```

This will test all Graph API endpoints and identify exactly which ones are failing.

### 2. Check Token Scopes
```bash
node decode-token.js
```

This will analyze the token permissions and identify missing scopes.

## 🔧 Step-by-Step Resolution

### Step 1: Verify Azure App Registration Configuration

1. **Go to Azure Portal**: https://portal.azure.com
2. **Navigate to**: Azure Active Directory → App registrations
3. **Find your app**: Look for the app with your `GRAPH_CLIENT_ID`
4. **Check Overview**: Verify Client ID and Tenant ID match your `.env` file

### Step 2: Verify API Permissions

1. **In your app registration, go to**: API permissions
2. **Ensure these delegated permissions are present**:
   - ✅ `Calendars.ReadWrite` (not Calendar.ReadWrite)
   - ✅ `User.Read`
   - ✅ `MailboxSettings.Read`
   - ✅ `Calendars.ReadWrite.Shared`
   - ✅ `OnlineMeetings.ReadWrite`

3. **Check permission status**: All should show "Granted for [Your Organization]"

### Step 3: Grant Admin Consent

1. **In API permissions section**: Look for "Grant admin consent for [Your Organization]"
2. **Click the button**: This grants consent for all users in your tenant
3. **Verify status**: All permissions should now show "Granted for [Your Organization]"

### Step 4: Check Authentication Configuration

1. **Go to**: Authentication in your app registration
2. **Verify platform configuration**: Should include "Mobile and desktop applications"
3. **Check enabled flows**: Device code flow should be enabled
4. **Verify redirect URIs**: Should include `https://login.microsoftonline.com/common/oauth2/nativeclient`

### Step 5: Test with Microsoft Graph Explorer

1. **Go to**: https://developer.microsoft.com/en-us/graph/graph-explorer
2. **Sign in**: Use the same account as your MCP server
3. **Test these endpoints**:
   ```
   GET /me
   GET /me/calendars
   GET /users/{email}/calendars
   GET /users/{email}/calendarView
   POST /users/{email}/events
   ```

4. **Compare results**: If these work in Graph Explorer but fail in your app, it's a permission issue

## 🚨 Common Issues and Solutions

### Issue 1: "No scopes found in token response"
**Symptoms**: Token is generated but has no scopes
**Solution**: 
- Check that delegated permissions are granted (not application permissions)
- Ensure admin consent has been given
- Verify the app supports your account type

### Issue 2: "401 Unauthorized" on calendar endpoints
**Symptoms**: Authentication works but calendar operations fail
**Solution**:
- Verify `Calendars.ReadWrite` permission is granted
- Check that admin consent includes calendar permissions
- Ensure your account has calendar access in the organization

### Issue 3: "403 Forbidden" on specific operations
**Symptoms**: Some operations work, others fail
**Solution**:
- Check if your organization has conditional access policies
- Verify the specific permission for the failing operation
- Contact your IT admin if needed

### Issue 4: Token expires quickly or authentication loops
**Symptoms**: Frequent re-authentication required
**Solution**:
- Check token expiration settings in Azure
- Verify the app registration supports long-lived tokens
- Check if your organization has token lifetime policies

## 🔍 Advanced Troubleshooting

### Check Token Claims
If you have access to the raw JWT token, you can decode it at https://jwt.ms to see:
- `scp` (scopes) - what permissions the token actually has
- `aud` (audience) - should be "https://graph.microsoft.com"
- `iss` (issuer) - should be from Microsoft
- `exp` (expiration) - when the token expires

### Check Organization Policies
Your organization may have:
- **Conditional Access**: Requiring specific devices, locations, or MFA
- **Token Lifetime Policies**: Limiting how long tokens are valid
- **API Restrictions**: Blocking certain Graph API endpoints
- **User Consent Policies**: Requiring admin approval for all apps

### Check Account Type
- **Work/School accounts**: Require admin consent for delegated permissions
- **Personal Microsoft accounts**: May have different permission models
- **Guest accounts**: May have limited access to resources

## 📋 Verification Checklist

- [ ] Azure app registration exists and is active
- [ ] Client ID and Tenant ID match your `.env` file
- [ ] Delegated permissions are granted (not application permissions)
- [ ] Admin consent has been given for all permissions
- [ ] Authentication configuration supports device code flow
- [ ] Redirect URIs include the required endpoints
- [ ] Your account is in the same Azure tenant
- [ ] Your account has the required permissions
- [ ] No conditional access policies are blocking access
- [ ] Graph Explorer can access the same endpoints

## 🆘 Still Having Issues?

If you've completed all the steps above and still get 401 errors:

1. **Run the diagnostic tools** to get detailed error information
2. **Check the logs** for specific error codes and messages
3. **Test with Graph Explorer** to isolate the issue
4. **Contact your IT admin** if it's an organization policy issue
5. **Check Microsoft Graph API status** for any service issues

## 📚 Additional Resources

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [Azure App Registration Guide](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- [Graph API Permissions Reference](https://docs.microsoft.com/en-us/graph/permissions-reference)
- [Troubleshooting Authentication](https://docs.microsoft.com/en-us/azure/active-directory/develop/authentication-vs-authorization)
