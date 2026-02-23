# Salesforce (Delphi) Phase 2 Integration Outline

## Recommended auth model
1. Use Salesforce Connected App (OAuth 2.0 Authorization Code flow).
2. Redirect user to Salesforce login URL.
3. Let Salesforce handle MFA/Authenticator OTP.
4. Exchange code for access token + refresh token.

## Integration tasks
1. Create Connected App in Salesforce with required scopes (`api`, `refresh_token`, `openid` if needed).
2. Add callback endpoint in app for token exchange.
3. Store encrypted tokens server-side (not user passwords).
4. Query Delphi/Salesforce objects for room inventory + availability by date range.
5. Add availability filter on top of current capacity ranking.

## Data needed from Delphi/Salesforce admin
- Object API names for meeting space inventory and availability
- Field API names for:
  - room name mapping
  - start/end date or date-time blocks
  - booked/unavailable flags
- API access policy and sandbox credentials for testing
