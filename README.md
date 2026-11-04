# MFA bypass risk evaluation and Microsoft 365 tenant hardening for SMEs

Fictional case: Coastal Managed IT-style SME on Microsoft 365 who “turned on MFA” and stopped there. This pack is **defensive only**.

## What this is

- High-level catalogue of MFA-bypass *classes* that threaten default-MFA tenants (token theft, consent phishing, legacy authentication gaps, weak MFA methods, session persistence).
- A proportionate hardening baseline and SME checklist.
- Pre/post residual-risk scoring for a synthetic laboratory baseline.
- Free-tier honesty: Conditional Access and Identity Protection need paid SKUs — Phase 0 uses what Free/security-defaults estates can actually do.

## What this is not

- No phishing-proxy tooling, no attack scripts, no token-steal procedures, no production-tenant testing.
- Offensive reproduction stays out of scope; risk is reasoned from public vendor/ACSC guidance and defensive control coverage.

## Research questions

1. Which MFA-bypass classes remain open against a default-MFA SME tenant?
2. What hardening closes those gaps affordably?
3. How much residual risk remains after the baseline?

## Run

```powershell
.\01_Score_Bypass_Exposure.ps1
.\02_Apply_Hardening_Checklist.ps1   # writes intended-state evidence (no live tenant changes by default)
.\03_Score_Residual_Risk.ps1
```
