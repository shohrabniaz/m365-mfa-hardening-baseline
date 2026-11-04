<#
.SYNOPSIS
  Record intended hardening baseline state (documentation evidence).
  Does not change a live tenant unless -LiveApply is used (still refuses unknown tenants).
#>
param([switch]$LiveApply)
. "$PSScriptRoot\00_Config.ps1"

if ($LiveApply) {
    Write-Warning "LiveApply refused in this pack by default policy. Document intended state only. Configure controls in a lab portal under change control."
}

$controls = @(
    [ordered]@{ Id = "H-01"; Title = "Block legacy authentication"; Phase = "1 (needs CA / equivalent)"; Closes = @("BYP-01"); Status = "Intended"; FreeTier = "NotAvailable — requires Conditional Access or Exchange auth policies where licensed" }
    [ordered]@{ Id = "H-02"; Title = "Restrict user consent to verified publishers / admin consent workflow"; Phase = "0/1"; Closes = @("BYP-02"); Status = "Intended"; FreeTier = "Partial — Entra enterprise app consent settings available on many SKUs" }
    [ordered]@{ Id = "H-03"; Title = "Disable users from authorizing new risky apps; review existing grants"; Phase = "0"; Closes = @("BYP-02"); Status = "Intended"; FreeTier = "Available — periodic admin review of enterprise applications" }
    [ordered]@{ Id = "H-04"; Title = "Shorten sign-in session / require reauth for sensitive apps"; Phase = "1"; Closes = @("BYP-03","BYP-05"); Status = "Intended"; FreeTier = "NotAvailable without CA sign-in frequency" }
    [ordered]@{ Id = "H-05"; Title = "Revoke sessions on high-risk leavers and suspected compromise"; Phase = "0"; Closes = @("BYP-03"); Status = "Intended"; FreeTier = "Available — revoke sessions / reset auth methods runbook" }
    [ordered]@{ Id = "H-06"; Title = "Prefer authenticator app / phishing-resistant methods for admins"; Phase = "0"; Closes = @("BYP-04"); Status = "Intended"; FreeTier = "Partial — method registration policies may need P1" }
    [ordered]@{ Id = "H-07"; Title = "Require MFA for all users (security defaults or CA)"; Phase = "0"; Closes = @("BYP-04"); Status = "Intended"; FreeTier = "Available — security defaults" }
    [ordered]@{ Id = "H-08"; Title = "Device compliance for admin and privileged roles"; Phase = "1"; Closes = @("BYP-03","BYP-05"); Status = "Intended"; FreeTier = "NotAvailable — Intune + CA" }
    [ordered]@{ Id = "H-09"; Title = "Two monitored break-glass accounts; alert on use"; Phase = "0"; Closes = @("BYP-06"); Status = "Intended"; FreeTier = "Available — process + mailbox/Teams alert" }
    [ordered]@{ Id = "H-10"; Title = "Weekly sign-in anomaly review (failed + unusual locations)"; Phase = "0"; Closes = @("BYP-03","BYP-05","BYP-06"); Status = "Intended"; FreeTier = "Available — Entra sign-in logs" }
)

$out = [ordered]@{
    Organisation = $Org
    Mode = "IntendedState"
    LiveChanges = $false
    Ethics = "Hardening checklist only. No offensive validation tooling."
    Controls = $controls
    Phase0Count = @($controls | Where-Object { $_.Phase -like "0*" }).Count
    Phase1Count = @($controls | Where-Object { $_.Phase -like "1*" }).Count
}
Write-EvidenceJson -Prefix "Hardening_Intended_State" -Object $out | Out-Null
Write-Host ("Recorded {0} hardening controls (intended state)." -f $controls.Count)
