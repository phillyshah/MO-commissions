-- A second, separately-hashed token per session for the sibling apps behind Traefik ForwardAuth.
-- The main cookie is __Host- (never leaves accounting.90ten.life). The SSO cookie is scoped to
-- Domain=accounting.90ten.life so usage.accounting.* and commissions.accounting.* can present it to
-- /auth/verify. It is still never scoped to .90ten.life, so MaxxDash and website cookies never reach us.
alter table portal_session add column sso_token_hash bytea unique;
