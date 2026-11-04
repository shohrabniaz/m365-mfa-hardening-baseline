"""Build MFA hardening / bypass-class evaluation submission pack (defensive only)."""
from __future__ import annotations

from datetime import date
from pathlib import Path
import json
import shutil

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from pptx import Presentation
from pptx.dml.color import RGBColor as PRgb
from pptx.util import Inches, Pt as PPt

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "Final_Submission"
DOCS = ROOT / "01_Assessment"
HARD = ROOT / "02_Hardening"
EVAL = ROOT / "03_Evaluation"
EVID = ROOT / "evidence"
NAVY = RGBColor(0x1F, 0x4E, 0x79)
TODAY = date.today().isoformat()


def heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for r in h.runs:
        r.font.color.rgb = NAVY


def load_json(pattern: str):
    files = sorted(EVID.glob(pattern), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        return None
    return json.loads(files[0].read_text(encoding="utf-8-sig"))


def build_docs(pre, post):
    DOCS.mkdir(parents=True, exist_ok=True)
    HARD.mkdir(parents=True, exist_ok=True)

    d = Document()
    d.add_heading("MFA bypass-class assessment (laboratory, defensive)", 0)
    d.add_paragraph(f"Md Shohrab Hossen Niaz | M365 MFA hardening | {TODAY}")
    heading(d, "Ethics boundary", 1)
    d.add_paragraph(
        "This assessment names publicly documented bypass CLASSES only "
        "(legacy auth gaps, illicit OAuth consent, post-MFA session theft, weak MFA methods, "
        "unmanaged persistent sessions, unmonitored break-glass). It does not include attack "
        "tooling, phishing-proxy setup, token-theft procedures, or tests against any live "
        "customer or employer tenant."
    )
    heading(d, "Pre-hardening exposure", 1)
    if pre:
        d.add_paragraph(
            f"Synthetic default-MFA exposure: {pre.get('AggregateExposurePercent')}% "
            f"({pre.get('RiskBand')})."
        )
        for c in pre.get("Classes", []):
            d.add_paragraph(
                f"{c.get('Id')} [{c.get('DefaultMfaExposure')}] {c.get('Class')} — {c.get('WhyOpen')}",
                style="List Bullet",
            )
    d.save(DOCS / "Bypass_Class_Assessment.docx")

    d = Document()
    d.add_heading("Tenant hardening configuration baseline", 0)
    d.add_paragraph(f"Md Shohrab Hossen Niaz | M365 MFA hardening | {TODAY}")
    heading(d, "Phase 0 (Free / security-defaults capable)", 1)
    for x in [
        "Enforce MFA for all interactive users (security defaults or equivalent).",
        "Restrict user OAuth consent; review existing enterprise app grants monthly.",
        "Revoke sessions runbook for leavers and suspected compromise.",
        "Prefer authenticator app for admins; discourage SMS-only for privileged roles.",
        "Two monitored break-glass accounts with use alerts.",
        "Weekly sign-in anomaly review.",
    ]:
        d.add_paragraph(x, style="List Number")
    heading(d, "Phase 1 (Conditional Access / Intune funded)", 1)
    for x in [
        "Block legacy authentication.",
        "Sign-in frequency and session controls for sensitive apps.",
        "Require compliant or hybrid-joined devices for admins.",
        "Optional risk-based controls when Identity Protection is licensed.",
    ]:
        d.add_paragraph(x, style="List Number")
    heading(d, "Free-tier honesty", 1)
    d.add_paragraph(
        "Do not claim Conditional Access or Identity Protection on Entra ID Free. "
        "Record NotAvailable and fund Phase 1 before promising partners those controls."
    )
    d.save(HARD / "Hardening_Configuration_Baseline.docx")

    d = Document()
    d.add_heading("SME Microsoft 365 MFA hardening checklist", 0)
    d.add_paragraph(f"Md Shohrab Hossen Niaz | {TODAY}")
    checks = [
        ("H-01", "Legacy auth blocked (or gap logged as Phase 1)"),
        ("H-02", "User consent restricted / admin consent workflow"),
        ("H-03", "Enterprise app grants reviewed in last 30 days"),
        ("H-04", "Session reauth policy for sensitive apps (Phase 1)"),
        ("H-05", "Session-revoke runbook tested once"),
        ("H-06", "Admins not on SMS-only MFA"),
        ("H-07", "MFA on for all interactive users"),
        ("H-08", "Admin device compliance (Phase 1)"),
        ("H-09", "Break-glass monitored"),
        ("H-10", "Weekly sign-in review owner named"),
    ]
    for cid, text in checks:
        d.add_paragraph(f"[ ] {cid} {text}", style="List Bullet")
    d.save(HARD / "SME_Hardening_Checklist.docx")

    d = Document()
    d.add_heading("Re-test and residual-risk report", 0)
    d.add_paragraph(f"Md Shohrab Hossen Niaz | {TODAY}")
    heading(d, "Method", 1)
    d.add_paragraph(
        "Residual risk is scored by control coverage against each bypass class after "
        "intended Phase 0 controls and designed Phase 1 controls. This is not a live "
        "adversarial re-test."
    )
    if post:
        heading(d, "Results", 1)
        d.add_paragraph(
            f"Pre-exposure {post.get('PreExposurePercent')}% → post {post.get('PostExposurePercent')}% "
            f"(reduction {post.get('AbsoluteReductionPoints')} points). Band: {post.get('RiskBand')}."
        )
        for r in post.get("Residual", []):
            d.add_paragraph(
                f"{r.get('Id')} residual {r.get('ResidualScore')}/{r.get('MaxScore')}: {r.get('Rationale')}",
                style="List Bullet",
            )
        heading(d, "Honest limits", 1)
        for lim in post.get("HonestLimits", []):
            d.add_paragraph(lim, style="List Bullet")
    d.save(DOCS / "Residual_Risk_Report.docx")


def build_workbook(pre, post):
    EVAL.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    fill = PatternFill("solid", fgColor="1F4E79")
    white = Font(color="FFFFFF", bold=True)

    dash = wb.active
    dash.title = "Risk_Reduction"
    dash["A1"] = "MFA bypass-class exposure — synthetic SME baseline"
    dash["A1"].font = Font(bold=True, color="1F4E79", size=13)
    dash["A3"] = "Pre-hardening %"
    dash["B3"] = pre.get("AggregateExposurePercent") if pre else ""
    dash["A4"] = "Post-hardening %"
    dash["B4"] = post.get("PostExposurePercent") if post else ""
    dash["A5"] = "Reduction (points)"
    dash["B5"] = post.get("AbsoluteReductionPoints") if post else ""
    dash["A6"] = "Risk band (post)"
    dash["B6"] = post.get("RiskBand") if post else ""
    dash.column_dimensions["A"].width = 28
    dash.column_dimensions["B"].width = 55

    pre_s = wb.create_sheet("Pre_Classes")
    pre_s.append(["Id", "Class", "Exposure", "Score", "Why open"])
    for c in pre_s[1]:
        c.fill = fill
        c.font = white
    if pre:
        for item in pre.get("Classes", []):
            pre_s.append(
                [
                    item.get("Id"),
                    item.get("Class"),
                    item.get("DefaultMfaExposure"),
                    item.get("ExposureScore"),
                    item.get("WhyOpen"),
                ]
            )
    for col in pre_s.columns:
        pre_s.column_dimensions[col[0].column_letter].width = 24
    pre_s.column_dimensions["B"].width = 42
    pre_s.column_dimensions["E"].width = 70

    post_s = wb.create_sheet("Residual")
    post_s.append(["Id", "Class", "ResidualScore", "Rationale"])
    for c in post_s[1]:
        c.fill = fill
        c.font = white
    if post:
        for item in post.get("Residual", []):
            post_s.append(
                [item.get("Id"), item.get("Class"), item.get("ResidualScore"), item.get("Rationale")]
            )
    for col in post_s.columns:
        post_s.column_dimensions[col[0].column_letter].width = 28
    post_s.column_dimensions["D"].width = 70

    chk = wb.create_sheet("Checklist")
    chk.append(["Id", "Control", "Phase", "Done (Y/N)", "Owner", "Date"])
    for c in chk[1]:
        c.fill = fill
        c.font = white
    for row in [
        ("H-01", "Block legacy auth", "1"),
        ("H-02", "Restrict OAuth consent", "0/1"),
        ("H-03", "Review app grants", "0"),
        ("H-04", "Sign-in frequency / reauth", "1"),
        ("H-05", "Session revoke runbook", "0"),
        ("H-06", "Strong MFA for admins", "0"),
        ("H-07", "MFA all users", "0"),
        ("H-08", "Admin device compliance", "1"),
        ("H-09", "Monitored break-glass", "0"),
        ("H-10", "Weekly sign-in review", "0"),
    ]:
        chk.append([row[0], row[1], row[2], "", "", ""])
    for col in chk.columns:
        chk.column_dimensions[col[0].column_letter].width = 26

    path = EVAL / "MFA_Hardening_Risk_Workbook.xlsx"
    wb.save(path)
    return path


def build_report(pre, post):
    pre_pct = pre.get("AggregateExposurePercent") if pre else "n/a"
    post_pct = post.get("PostExposurePercent") if post else "n/a"
    red = post.get("AbsoluteReductionPoints") if post else "n/a"
    band = post.get("RiskBand") if post else "n/a"

    sections = [
        (
            "Abstract",
            "SMEs often enable MFA on Microsoft 365 and assume the job is finished. "
            "Attackers still abuse legacy authentication gaps, illicit OAuth consent, "
            "post-MFA session theft, weak authenticator methods, unmanaged persistent "
            "sessions, and unmonitored break-glass accounts. I built a defensive "
            "assessment that scores those bypass classes on a synthetic default-MFA "
            f"tenant, designed a Phase 0/Phase 1 hardening baseline, and re-scored "
            f"residual risk. Pre-hardening exposure is {pre_pct}%; post-hardening "
            f"modelled residual is {post_pct}% (reduction {red} points; {band}). "
            "No offensive tooling or live bypass reproduction is included.",
        ),
        (
            "1. Introduction",
            "Catalogue Project 37 asks which MFA-bypass techniques succeed against a "
            "default-MFA SME tenant, what hardening closes the gaps affordably, and how "
            "much residual risk remains. Offensive elements must stay in an isolated lab "
            "with clear ethics. In this environment, Graph Conditional Access was not "
            "available and reproducing phishing-proxy or token-theft tradecraft would be "
            "unsafe and out of policy. The artefact therefore evaluates bypass CLASSES "
            "from public Microsoft, ACSC/ASD, NIST, and ATT&CK guidance, then maps "
            "proportionate hardening controls.",
        ),
        (
            "2. Method",
            "Design-science and control-coverage measurement. Six bypass classes were "
            "scored 0–100 for a synthetic default-MFA estate. A ten-control hardening "
            "baseline was written as intended state (Phase 0 Free/security-defaults "
            "capable; Phase 1 Conditional Access and Intune). Residual scores assume "
            "Phase 0 applied and Phase 1 designed. JSON evidence and an Excel workbook "
            "record pre/post figures. Live tenant mutation is refused by the scripts.",
        ),
        (
            "3. Pre-hardening results",
            f"Aggregate exposure {pre_pct}% falls in a high false-confidence band: MFA is "
            "present, yet legacy auth, consent phishing, and session reuse remain open. "
            "Weak MFA methods and unmanaged devices add medium exposure. Break-glass "
            "without monitoring is a process bypass. This matches the business problem — "
            "a false sense of safety after ticking MFA.",
        ),
        (
            "4. Hardening baseline",
            "Phase 0 prioritises consent restriction, grant review, session-revoke "
            "runbooks, stronger admin authenticators, universal MFA via security "
            "defaults, monitored break-glass, and weekly sign-in review. Phase 1 adds "
            "legacy-auth block, sign-in frequency, and device compliance when licences "
            "exist. The SME checklist is a one-page owner-dated control register.",
        ),
        (
            "5. Residual risk",
            f"Modelled residual {post_pct}% ({band}) with about {red} points of reduction. "
            "Consent and break-glass classes fall sharply. Legacy auth and unmanaged "
            "session classes remain elevated until Conditional Access and Intune are "
            "funded — recorded as NotAvailable on Free tier rather than claimed. "
            "Residual scoring is not a red-team proof.",
        ),
        (
            "6. Ethics and limits",
            "No Evilginx-class tooling, no phishing pages, no token-steal scripts, no "
            "employer or client tenant. Screenshots of paid portals are not invented. "
            "Students or operators must not point these ideas at production without "
            "change control and written authorisation.",
        ),
        (
            "7. Conclusion",
            "MFA is necessary and not sufficient. An SME can cut meaningful identity risk "
            "with Phase 0 hygiene, then buy Conditional Access to close legacy and "
            "device-bound gaps. The measurable story is exposure percent down and a "
            "checklist with named owners — not a claim of unbypassable MFA.",
        ),
        (
            "References",
            "ASD Essential Eight (MFA). ACSC Microsoft 365 hardening guidance. Microsoft "
            "Learn: block legacy authentication; manage user consent; Conditional Access; "
            "security defaults. NIST SP 800-63B. MITRE ATT&CK credential-access techniques "
            "(conceptual mapping only).",
        ),
    ]
    bw = sum(len(b.split()) for t, b in sections if t != "References")
    m = Document()
    tp = m.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = tp.add_run(
        "Evaluating MFA bypass risk and hardening\n"
        "SME Microsoft 365 tenants"
    )
    tr.bold = True
    tr.font.size = Pt(16)
    tr.font.color.rgb = NAVY
    mp = m.add_paragraph()
    mp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    mr = mp.add_run(
        f"Md Shohrab Hossen Niaz\nM365 MFA hardening baseline (defensive)\n"
        f"Synthetic lab model | No offensive tooling\n"
        f"Word count (excluding title block and reference list): {bw}"
    )
    mr.font.size = Pt(11)
    for title, body in sections:
        heading(m, title, 1)
        m.add_paragraph(body)
    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "01_MFA_Hardening_Report.docx"
    m.save(path)
    print("BODY_WORDS", bw)
    return path


def build_slides(pre, post):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    navy = PRgb(0x1F, 0x4E, 0x79)

    def add(title, bullets, notes):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        b = s.shapes.add_textbox(Inches(0.6), Inches(0.4), Inches(12), Inches(1))
        r = b.text_frame.paragraphs[0].add_run()
        r.text = title
        r.font.size = PPt(28)
        r.font.bold = True
        r.font.color.rgb = navy
        body = s.shapes.add_textbox(Inches(0.6), Inches(1.5), Inches(12), Inches(5))
        tf = body.text_frame
        tf.word_wrap = True
        first = True
        for line in bullets:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            run = p.add_run()
            run.text = line
            run.font.size = PPt(18)
            p.space_after = PPt(10)
        s.notes_slide.notes_text_frame.text = notes

    pre_pct = pre.get("AggregateExposurePercent") if pre else "n/a"
    post_pct = post.get("PostExposurePercent") if post else "n/a"
    red = post.get("AbsoluteReductionPoints") if post else "n/a"

    add(
        "MFA is on — are you safe?",
        ["Default-MFA SME tenants still leave bypass classes open", "Defensive assessment only — no attack tooling", "Phase 0 hygiene + Phase 1 Conditional Access"],
        "Project 37 style hardening study.",
    )
    add(
        "Bypass classes (high level)",
        ["Legacy auth gaps", "Illicit OAuth consent", "Post-MFA session theft", "Weak methods, unmanaged sessions, break-glass"],
        "Classes not recipes.",
    )
    add(
        "Hardening and residual risk",
        [f"Exposure {pre_pct}% → {post_pct}% (≈{red} points down)", "Checklist with owners", "Free-tier: do not claim CA / ID Protection"],
        "Results.",
    )
    add(
        "Takeaway",
        ["MFA necessary, not sufficient", "Fund CA to close legacy + device gaps", "Ethics: lab/synthetic only — no live customer tests"],
        "Questions?",
    )
    path = OUT / "02_MFA_Hardening_Presentation.pptx"
    prs.save(path)
    return path


def assemble():
    pre = load_json("*_Bypass_Exposure_Pre.json")
    post = load_json("*_Residual_Risk_Post.json")
    build_docs(pre, post)
    build_workbook(pre, post)
    build_report(pre, post)
    build_slides(pre, post)

    pairs = [
        (EVAL / "MFA_Hardening_Risk_Workbook.xlsx", OUT / "Appendix" / "A_Evaluation" / "MFA_Hardening_Risk_Workbook.xlsx"),
        (DOCS / "Bypass_Class_Assessment.docx", OUT / "Appendix" / "B_Assessment" / "Bypass_Class_Assessment.docx"),
        (DOCS / "Residual_Risk_Report.docx", OUT / "Appendix" / "B_Assessment" / "Residual_Risk_Report.docx"),
        (HARD / "Hardening_Configuration_Baseline.docx", OUT / "Appendix" / "C_Hardening" / "Hardening_Configuration_Baseline.docx"),
        (HARD / "SME_Hardening_Checklist.docx", OUT / "Appendix" / "C_Hardening" / "SME_Hardening_Checklist.docx"),
    ]
    for src, dst in pairs:
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
    app_d = OUT / "Appendix" / "D_Scripts"
    app_e = OUT / "Appendix" / "E_Evidence"
    app_d.mkdir(parents=True, exist_ok=True)
    app_e.mkdir(parents=True, exist_ok=True)
    for f in ROOT.glob("*.ps1"):
        if not f.name.startswith("_"):
            shutil.copy2(f, app_d / f.name)
    for f in EVID.glob("*.json"):
        shutil.copy2(f, app_e / f.name)

    readme = Document()
    readme.add_heading("Project submission", 0)
    readme.add_paragraph(
        "Md Shohrab Hossen Niaz\nM365 MFA hardening baseline\n"
        "Defensive bypass-class assessment + hardening checklist. No offensive tooling."
    )
    readme.save(OUT / "00_Read_Me_First.docx")
    print("assembled", OUT)


if __name__ == "__main__":
    assemble()
