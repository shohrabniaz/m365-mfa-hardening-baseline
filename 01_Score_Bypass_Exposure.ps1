<#
.SYNOPSIS
  Score MFA-bypass class exposure for a synthetic default-MFA SME tenant.
  Defensive assessment only — no attack tooling.
#>
. "$PSScriptRoot\00_Config.ps1"

# Classes are named at ATT&CK / vendor-guidance level. No reproduction steps.
$classes = @(
    [ordered]@{
        Id = "BYP-01"; Class = "Legacy / basic authentication paths"
        AttckHint = "Credential Access — valid accounts via legacy protocols"
        DefaultMfaExposure = "High"
        WhyOpen = "MFA often does not apply to legacy mail protocols; many SMEs leave them enabled."
        ExposureScore = 85; MaxScore = 100
    },
    [ordered]@{
        Id = "BYP-02"; Class = "OAuth consent phishing / illicit grant"
        AttckHint = "Initial Access / Persistence — application consent"
        DefaultMfaExposure = "High"
        WhyOpen = "User can approve a malicious app after MFA; tenant consent policies often unrestricted."
        ExposureScore = 80; MaxScore = 100
    },
    [ordered]@{
        Id = "BYP-03"; Class = "Session / refresh-token theft (post-MFA)"
        AttckHint = "Credential Access — steal application access token"
        DefaultMfaExposure = "High"
        WhyOpen = "MFA protects interactive sign-in, not a stolen refresh token until revoked or Conditional Access binds device/risk."
        ExposureScore = 75; MaxScore = 100
    },
    [ordered]@{
        Id = "BYP-04"; Class = "Weak MFA methods (SMS / voice / push fatigue)"
        AttckHint = "Credential Access — MFA fatigue / SIM-related risks (high-level)"
        DefaultMfaExposure = "Medium"
        WhyOpen = "Default MFA methods may be phishing-susceptible; phishing-resistant methods not required."
        ExposureScore = 60; MaxScore = 100
    },
    [ordered]@{
        Id = "BYP-05"; Class = "Persistent browser sessions on unmanaged devices"
        AttckHint = "Defense Evasion / Credential Access — session reuse"
        DefaultMfaExposure = "Medium"
        WhyOpen = "No device compliance or sign-in frequency controls on Free tier / defaults."
        ExposureScore = 55; MaxScore = 100
    },
    [ordered]@{
        Id = "BYP-06"; Class = "Break-glass / excluded accounts without monitoring"
        AttckHint = "Privilege Escalation — valid accounts"
        DefaultMfaExposure = "Medium"
        WhyOpen = "Emergency accounts excluded from MFA with weak monitoring create a standing bypass."
        ExposureScore = 50; MaxScore = 100
    }
)

$total = ($classes | ForEach-Object { $_.ExposureScore } | Measure-Object -Sum).Sum
$max = ($classes | ForEach-Object { $_.MaxScore } | Measure-Object -Sum).Sum
$pct = [math]::Round(100.0 * $total / $max, 1)

$out = [ordered]@{
    Organisation = $Org
    Mode = "Synthetic"
    Ethics = "Assessment names bypass CLASSES only. No tooling, payloads, or reproduction procedures."
    Phase = "Pre-hardening"
    AggregateExposurePercent = $pct
    RiskBand = if ($pct -ge 70) { "High false-confidence" } elseif ($pct -ge 40) { "Elevated" } else { "Moderate" }
    Classes = $classes
    GuidanceRefs = @(
        "ACSC / ASD Essential Eight — Multi-factor authentication",
        "Microsoft: Block legacy authentication; Conditional Access",
        "Microsoft: Managing consent to applications",
        "NIST SP 800-63B — authenticator assurance",
        "MITRE ATT&CK — credential access (conceptual mapping only)"
    )
}
Write-EvidenceJson -Prefix "Bypass_Exposure_Pre" -Object $out | Out-Null
Write-Host ("Pre-hardening exposure {0}% — {1}" -f $pct, $out.RiskBand)
