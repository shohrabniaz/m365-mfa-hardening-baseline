<#
.SYNOPSIS
  Score residual exposure after intended Phase 0 + designed Phase 1 controls.
#>
. "$PSScriptRoot\00_Config.ps1"

# Residual scores assume Phase 0 applied + Phase 1 designed (not all live on Free tier).
$residual = @(
    [ordered]@{ Id = "BYP-01"; Class = "Legacy / basic authentication paths"; ResidualScore = 35; MaxScore = 100; Rationale = "Closed only when CA/auth policies exist (Phase 1). Phase 0 documents the gap." }
    [ordered]@{ Id = "BYP-02"; Class = "OAuth consent phishing / illicit grant"; ResidualScore = 25; MaxScore = 100; Rationale = "Consent restriction + grant review materially reduce this class." }
    [ordered]@{ Id = "BYP-03"; Class = "Session / refresh-token theft (post-MFA)"; ResidualScore = 40; MaxScore = 100; Rationale = "Revoke runbook helps; device-bound CA needed for stronger reduction." }
    [ordered]@{ Id = "BYP-04"; Class = "Weak MFA methods"; ResidualScore = 30; MaxScore = 100; Rationale = "Security defaults + stronger methods for admins; phishing-resistant fleet is later." }
    [ordered]@{ Id = "BYP-05"; Class = "Persistent sessions on unmanaged devices"; ResidualScore = 45; MaxScore = 100; Rationale = "Residual until Intune+CA; mitigated partly by reauth guidance." }
    [ordered]@{ Id = "BYP-06"; Class = "Break-glass without monitoring"; ResidualScore = 15; MaxScore = 100; Rationale = "Monitored break-glass + alerts largely close this process gap." }
)

$pre = 67.5  # approximate from equal-weight prior model; overwritten if pre file exists
$preFiles = Get-ChildItem $EvidenceDir -Filter "*_Bypass_Exposure_Pre.json" | Sort-Object LastWriteTime -Descending
if ($preFiles) {
    $preObj = Get-Content $preFiles[0].FullName -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($preObj.AggregateExposurePercent) { $pre = [double]$preObj.AggregateExposurePercent }
}

$total = ($residual | ForEach-Object { $_.ResidualScore } | Measure-Object -Sum).Sum
$max = ($residual | ForEach-Object { $_.MaxScore } | Measure-Object -Sum).Sum
$pct = [math]::Round(100.0 * $total / $max, 1)
$reduction = [math]::Round($pre - $pct, 1)

$out = [ordered]@{
    Organisation = $Org
    Mode = "SyntheticResidual"
    PreExposurePercent = $pre
    PostExposurePercent = $pct
    AbsoluteReductionPoints = $reduction
    RiskBand = if ($pct -ge 70) { "High false-confidence" } elseif ($pct -ge 40) { "Elevated" } else { "Moderate — continue Phase 1 funding" }
    Residual = $residual
    HonestLimits = @(
        "Free-tier lab cannot fully close BYP-01/BYP-03/BYP-05 without Conditional Access and device compliance.",
        "No live adversarial re-test was performed; residual is control-coverage scoring, not red-team proof.",
        "Do not claim phishing-resistant MFA fleet maturity without method registration evidence."
    )
}
Write-EvidenceJson -Prefix "Residual_Risk_Post" -Object $out | Out-Null
Write-Host ("Residual {0}% (was {1}%); reduction {2} points — {3}" -f $pct, $pre, $reduction, $out.RiskBand)
