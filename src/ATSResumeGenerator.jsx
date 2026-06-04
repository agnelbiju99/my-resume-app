/**
 * ATS Resume Generator
 * A fully functional React component for building ATS-optimized resumes.
 *
 * SETUP INSTRUCTIONS:
 * 1. Install dependencies:
 *    npm install docx file-saver
 *
 * 2. Import in your app:
 *    import ATSResumeGenerator from './ATSResumeGenerator';
 *
 * 3. The component uses only React + docx + file-saver (no UI library needed).
 *    Styles are all inline / CSS-in-JS for portability.
 *
 * FEATURES:
 * - 6-step form wizard (Personal → Summary → Experience → Education → Skills → Optional)
 * - Paste & Parse mode (uses Anthropic API to extract resume sections)
 * - Live resume preview (4 templates: Classic, Executive, Modern, Creative)
 * - Real-time ATS score (0–100) with checklist
 * - Job Description keyword analyzer with match scoring
 * - Download as .docx Word file (fully formatted, ATS-safe)
 * - Download as HTML
 * - Print to PDF
 * - Version manager (save up to 5 named versions)
 * - Profession presets with keyword chips and action verbs
 */

import { useState, useCallback, useEffect, useRef } from "react";
import { Document, Packer, Paragraph, TextRun, AlignmentType, LevelFormat, BorderStyle } from "docx";
import { saveAs } from "file-saver";

// ─── DATA ────────────────────────────────────────────────────────────────────

const PROFESSIONS = [
  { id: "tech", label: "Tech & Engineering", icon: "💻" },
  { id: "business", label: "Business & Finance", icon: "💼" },
  { id: "healthcare", label: "Healthcare & Medical", icon: "🏥" },
  { id: "creative", label: "Creative & Design", icon: "🎨" },
  { id: "legal", label: "Legal & Law", icon: "⚖️" },
  { id: "marketing", label: "Marketing & Sales", icon: "📊" },
  { id: "education", label: "Education & Research", icon: "🎓" },
  { id: "operations", label: "Operations & Logistics", icon: "🏗️" },
];

const KW_PRESETS = {
  tech: ["JavaScript","Python","React","Node.js","AWS","Docker","Kubernetes","CI/CD","REST API","SQL","Agile","Scrum","Git","TypeScript","Machine Learning","Cloud Architecture","Microservices","DevOps"],
  business: ["Financial Analysis","P&L Management","Budget Forecasting","Stakeholder Management","Strategic Planning","Business Development","ROI","KPI","Excel","Power BI","Risk Management","Compliance","M&A"],
  healthcare: ["Patient Care","Clinical Assessment","EHR","HIPAA Compliance","Evidence-Based Practice","Care Coordination","CPR/BLS","Medical Terminology","Quality Improvement","Patient Outcomes"],
  creative: ["UI/UX Design","Figma","Adobe Creative Suite","Brand Identity","Wireframing","Prototyping","User Research","Responsive Design","Typography","Visual Communication","Motion Graphics"],
  legal: ["Legal Research","Contract Drafting","Litigation","Due Diligence","Regulatory Compliance","Client Counseling","Legal Writing","Westlaw","LexisNexis","Negotiation","Case Management"],
  marketing: ["Digital Marketing","SEO/SEM","Content Strategy","Lead Generation","CRM","Salesforce","Campaign Management","Google Analytics","Conversion Optimization","Brand Strategy","Social Media"],
  education: ["Curriculum Development","Student Assessment","Instructional Design","Research Methods","Grant Writing","Classroom Management","Learning Outcomes","EdTech","Academic Writing","Mentoring"],
  operations: ["Supply Chain Management","Logistics Coordination","Process Improvement","Lean Six Sigma","Vendor Management","Inventory Control","ERP Systems","Project Management","KPI Tracking","Cost Reduction"],
};

const ACTION_VERBS = {
  tech: ["Architected","Deployed","Optimized","Engineered","Automated","Migrated","Implemented","Refactored"],
  business: ["Spearheaded","Negotiated","Forecasted","Streamlined","Orchestrated","Leveraged","Drove","Aligned"],
  healthcare: ["Administered","Assessed","Coordinated","Documented","Monitored","Facilitated","Implemented","Educated"],
  creative: ["Designed","Conceptualized","Produced","Crafted","Visualized","Rebranded","Illustrated","Prototyped"],
  legal: ["Drafted","Negotiated","Litigated","Counseled","Researched","Reviewed","Advised","Argued"],
  marketing: ["Launched","Grew","Converted","Targeted","Analyzed","Executed","Generated","Optimized"],
  education: ["Developed","Facilitated","Mentored","Assessed","Published","Designed","Instructed","Researched"],
  operations: ["Optimized","Reduced","Managed","Coordinated","Streamlined","Implemented","Oversaw","Delivered"],
};

const TEMPLATES = [
  { id: "classic",   name: "Professional Classic", sub: "Times New Roman, traditional" },
  { id: "executive", name: "Executive",             sub: "Clean serif, blue headers" },
  { id: "modern",    name: "Modern Minimal",        sub: "Sans-serif, color bar" },
  { id: "creative",  name: "Creative Clean",        sub: "Accent header" },
];

const STEPS = [
  { id: "personal",   label: "Personal" },
  { id: "summary",    label: "Summary" },
  { id: "experience", label: "Experience" },
  { id: "education",  label: "Education" },
  { id: "skills",     label: "Skills" },
  { id: "optional",   label: "Optional" },
];

const emptyData = () => ({
  name: "", email: "", phone: "", location: "", linkedin: "", website: "",
  summary: "",
  experience: [{ id: 1, title: "", company: "", location: "", startDate: "", endDate: "", bullets: [""] }],
  education: [{ id: 1, degree: "", institution: "", year: "", gpa: "", honors: "", university: "", mbbsClass: "" }],
  skills: { technical: [], soft: [], tools: [] },
  skillInput: { technical: "", soft: "", tools: "" },
  certifications: "",
  optionals: { projects: false, publications: false, volunteer: false, awards: false, languages: false },
  projects: [], volunteer: "", awards: "", languages: "", publications: "",
});

// ─── WORD EXPORT ──────────────────────────────────────────────────────────────

async function exportToWord(data) {
  const allSkills = [...data.skills.technical, ...data.skills.soft, ...data.skills.tools];

  const sectionTitle = (text) =>
    new Paragraph({
      children: [new TextRun({ text, bold: true, size: 22, font: "Times New Roman" })],
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "333333", space: 1 } },
      spacing: { before: 200, after: 80 },
    });

  const contactLine = [data.email, data.phone, data.location, data.linkedin, data.website]
    .filter(Boolean).join("  |  ");

  const children = [
    // Name
    new Paragraph({
      children: [new TextRun({ text: data.name || "Your Name", bold: true, size: 36, font: "Times New Roman" })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 60 },
    }),
    // Contact
    new Paragraph({
      children: [new TextRun({ text: contactLine, size: 18, font: "Times New Roman", color: "444444" })],
      alignment: AlignmentType.CENTER,
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "999999", space: 1 } },
      spacing: { after: 160 },
    }),
  ];

  // Summary
  if (data.summary) {
    children.push(sectionTitle("PROFESSIONAL SUMMARY"));
    children.push(new Paragraph({
      children: [new TextRun({ text: data.summary, size: 20, font: "Times New Roman" })],
      spacing: { after: 120 },
    }));
  }

  // Experience
  const expEntries = data.experience.filter(e => e.title);
  if (expEntries.length) {
    children.push(sectionTitle("WORK EXPERIENCE"));
    expEntries.forEach(e => {
      const dateStr = [e.startDate, e.endDate].filter(Boolean).join(" – ");
      children.push(new Paragraph({
        children: [
          new TextRun({ text: e.title + (e.company ? ", " + e.company : ""), bold: true, size: 20, font: "Times New Roman" }),
          new TextRun({ text: dateStr ? "  " + dateStr : "", size: 20, font: "Times New Roman", color: "555555" }),
        ],
        spacing: { before: 100, after: 40 },
      }));
      if (e.location) {
        children.push(new Paragraph({
          children: [new TextRun({ text: e.location, size: 18, font: "Times New Roman", color: "666666", italics: true })],
          spacing: { after: 40 },
        }));
      }
      e.bullets.filter(b => b.trim()).forEach(b => {
        children.push(new Paragraph({
          numbering: { reference: "resume-bullets", level: 0 },
          children: [new TextRun({ text: b.replace(/^[-•]\s*/, ""), size: 19, font: "Times New Roman" })],
          spacing: { after: 30 },
        }));
      });
    });
  }

  // Education
  const eduEntries = data.education.filter(e => e.degree);
  if (eduEntries.length) {
    children.push(sectionTitle("EDUCATION"));
    eduEntries.forEach(e => {
      children.push(new Paragraph({
        children: [
          new TextRun({ text: e.degree, bold: true, size: 20, font: "Times New Roman" }),
          new TextRun({ text: e.year ? "  " + e.year : "", size: 20, font: "Times New Roman", color: "555555" }),
        ],
        spacing: { before: 100, after: 40 },
      }));
      const isBtech = /b\.?tech|b\.?e\b/i.test(e.degree); const isMbbs = /mbbs/i.test(e.degree); const sub = [e.institution, e.university||"", e.gpa ? (isBtech ? "CGPA: " : "Percentage: ") + e.gpa : "", e.mbbsClass ? "Class: "+e.mbbsClass : "", e.honors].filter(Boolean).join("  |  ");
      if (sub) {
        children.push(new Paragraph({
          children: [new TextRun({ text: sub, size: 18, font: "Times New Roman", color: "555555" })],
          spacing: { after: 60 },
        }));
      }
    });
  }

  // Certifications
  if (data.certifications) {
    children.push(sectionTitle("CERTIFICATIONS"));
    data.certifications.split("\n").filter(Boolean).forEach(c => {
      children.push(new Paragraph({
        numbering: { reference: "resume-bullets", level: 0 },
        children: [new TextRun({ text: c, size: 19, font: "Times New Roman" })],
        spacing: { after: 30 },
      }));
    });
  }

  // Skills
  if (allSkills.length) {
    children.push(sectionTitle("SKILLS"));
    const skillGroups = [];
    if (data.skills.technical.length) skillGroups.push("Technical: " + data.skills.technical.join(", "));
    if (data.skills.tools.length) skillGroups.push("Tools & Software: " + data.skills.tools.join(", "));
    if (data.skills.soft.length) skillGroups.push("Soft Skills: " + data.skills.soft.join(", "));
    skillGroups.forEach(g => {
      children.push(new Paragraph({
        children: [new TextRun({ text: g, size: 19, font: "Times New Roman" })],
        spacing: { after: 50 },
      }));
    });
  }

  // Optional sections
  if (data.optionals.projects && Array.isArray(data.projects) && data.projects.length) {
    children.push(sectionTitle("PROJECTS"));
    data.projects.filter(Boolean).forEach(p => {
      const [name, desc, url] = p.split("|").map(s => s.trim());
      children.push(new Paragraph({
        children: [
          new TextRun({ text: name || "", bold: true, size: 20, font: "Times New Roman" }),
          desc ? new TextRun({ text: "  – " + desc + (url ? "  " + url : ""), size: 19, font: "Times New Roman" }) : new TextRun(""),
        ],
        spacing: { before: 80, after: 40 },
      }));
    });
  }

  if (data.optionals.awards && data.awards) {
    children.push(sectionTitle("AWARDS & HONORS"));
    data.awards.split("\n").filter(Boolean).forEach(a => {
      children.push(new Paragraph({
        numbering: { reference: "resume-bullets", level: 0 },
        children: [new TextRun({ text: a, size: 19, font: "Times New Roman" })],
        spacing: { after: 30 },
      }));
    });
  }

  if (data.optionals.languages && data.languages) {
    children.push(sectionTitle("LANGUAGES"));
    data.languages.split("\n").filter(Boolean).forEach(l => {
      children.push(new Paragraph({
        numbering: { reference: "resume-bullets", level: 0 },
        children: [new TextRun({ text: l, size: 19, font: "Times New Roman" })],
        spacing: { after: 30 },
      }));
    });
  }

  if (data.optionals.volunteer && data.volunteer) {
    children.push(sectionTitle("VOLUNTEER WORK"));
    children.push(new Paragraph({
      children: [new TextRun({ text: data.volunteer, size: 19, font: "Times New Roman" })],
      spacing: { after: 80 },
    }));
  }

  if (data.optionals.publications && data.publications) {
    children.push(sectionTitle("PUBLICATIONS"));
    data.publications.split("\n").filter(Boolean).forEach(p => {
      children.push(new Paragraph({
        numbering: { reference: "resume-bullets", level: 0 },
        children: [new TextRun({ text: p, size: 19, font: "Times New Roman" })],
        spacing: { after: 30 },
      }));
    });
  }

  const doc = new Document({
    numbering: {
      config: [{
        reference: "resume-bullets",
        levels: [{
          level: 0,
          format: LevelFormat.BULLET,
          text: "–",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 440, hanging: 220 } } },
        }],
      }],
    },
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
        },
      },
      children,
    }],
  });

  // Use toBlob for browser compatibility (avoids Node.js Buffer issue)
  const blob = await Packer.toBlob(doc);
  saveAs(blob, (data.name || "resume").replace(/\s+/g, "_") + "_resume.docx");
}

// ─── SCORE ENGINE ─────────────────────────────────────────────────────────────

function calcScore(data, prof) {
  let score = 0;
  const checks = [];
  const contactFull = data.name && data.email && data.phone && data.location;
  if (contactFull) { score += 10; checks.push({ pass: true, text: "Contact info complete (+10)" }); }
  else checks.push({ pass: false, text: "Complete name, email, phone, location" });

  const wc = data.summary.split(/\s+/).filter(w => w).length;
  if (wc >= 40 && wc <= 80) { score += 15; checks.push({ pass: true, text: `Summary ${wc} words, perfect (+15)` }); }
  else if (wc > 0) checks.push({ pass: false, text: `Summary: ${wc} words (need 40–80)` });
  else checks.push({ pass: false, text: "Add a professional summary" });

  const expWithBullets = data.experience.filter(e => e.title && e.bullets.some(b => b.trim())).length;
  if (expWithBullets >= 2) { score += 20; checks.push({ pass: true, text: "2+ experience entries with bullets (+20)" }); }
  else checks.push({ pass: false, text: `Add ${2 - expWithBullets} more experience entries with bullets` });

  const allSkills = [...data.skills.technical, ...data.skills.soft, ...data.skills.tools];
  if (allSkills.length >= 6) { score += 15; checks.push({ pass: true, text: `${allSkills.length} skills listed (+15)` }); }
  else checks.push({ pass: false, text: `Add ${6 - allSkills.length} more skills (${allSkills.length}/6 minimum)` });

  if (data.education.some(e => e.degree && e.institution)) { score += 10; checks.push({ pass: true, text: "Education section complete (+10)" }); }
  else checks.push({ pass: false, text: "Complete your education section" });

  const verbs = ACTION_VERBS[prof] || [];
  const allBullets = data.experience.flatMap(e => e.bullets).join(" ").toLowerCase();
  const hasVerbs = verbs.some(v => allBullets.includes(v.toLowerCase()));
  if (hasVerbs) { score += 10; checks.push({ pass: true, text: "Strong action verbs detected (+10)" }); }
  else checks.push({ pass: false, text: `Use action verbs: ${verbs.slice(0, 3).join(", ")}...` });

  score += 10;
  checks.push({ pass: true, text: "ATS-safe single-column formatting (+10)" });

  const profKws = KW_PRESETS[prof] || [];
  const resumeText = (data.summary + " " + allSkills.join(" ")).toLowerCase();
  const kwMatch = profKws.filter(k => resumeText.includes(k.toLowerCase())).length;
  if (kwMatch >= 3) { score += 10; checks.push({ pass: true, text: `${kwMatch} profession keywords present (+10)` }); }
  else checks.push({ pass: false, text: `Add more ${PROFESSIONS.find(p => p.id === prof)?.label} keywords` });

  return { score, checks };
}

// ─── STYLES ───────────────────────────────────────────────────────────────────

const S = {
  app: { fontFamily: "'Segoe UI', system-ui, sans-serif", fontSize: 14, color: "#1a1a2e", minHeight: "100vh", background: "#f0f2f5", display: "flex", flexDirection: "column" },
  topbar: { background: "#0f1629", color: "#fff", padding: "10px 14px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 6, flexWrap: "wrap" },
  appTitle: { fontWeight: 600, fontSize: 15, letterSpacing: ".2px", display: "flex", alignItems: "center", gap: 7 },
  modeTabs: { display: "flex", borderRadius: 8, overflow: "hidden", border: "1px solid #2d3a5a" },
  modeTab: (active) => ({ padding: "5px 14px", fontSize: 12, cursor: "pointer", background: active ? "#2563eb" : "transparent", color: active ? "#fff" : "#8899bb", border: "none", fontFamily: "inherit", transition: "all .15s" }),
  btnGroup: { display: "flex", gap: 6, alignItems: "center" },
  btn: { padding: "6px 12px", borderRadius: 7, border: "1px solid #dde1e9", background: "#fff", color: "#1a1a2e", fontSize: 12, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 5, fontFamily: "inherit", transition: "all .15s" },
  btnPrimary: { background: "#2563eb", color: "#fff", border: "1px solid #2563eb" },
  btnWord: { background: "#1d6f42", color: "#fff", border: "1px solid #1d6f42" },
  btnSm: { padding: "4px 9px", fontSize: 11 },
  main: { display: "flex", flex: 1, overflow: "hidden" },
  mobileOnly: { display: "none" },
  editorCol: { flex: 1, minWidth: 0, overflowY: "auto", padding: 16, display: "flex", flexDirection: "column", gap: 12 },
  previewCol: { width: 390, minWidth: 320, background: "#e8ebf0", borderLeft: "1px solid #d0d5de", overflowY: "auto", padding: 16, display: "flex", flexDirection: "column", gap: 12 },
  card: { background: "#fff", border: "1px solid #e5e7ef", borderRadius: 12, padding: 14 },
  cardTitle: { fontSize: 13, fontWeight: 600, marginBottom: 10, display: "flex", alignItems: "center", gap: 6, color: "#1a1a2e" },
  formRow: (cols = 2) => ({ display: "grid", gridTemplateColumns: `repeat(${cols}, 1fr)`, gap: 8, marginBottom: 8 }),
  formGroup: { display: "flex", flexDirection: "column", gap: 3 },
  label: { fontSize: 10, color: "#6b7280", fontWeight: 600, textTransform: "uppercase", letterSpacing: ".5px" },
  input: { fontSize: 13, fontFamily: "inherit", padding: "7px 10px", border: "1px solid #d1d5db", borderRadius: 7, background: "#fff", color: "#1a1a2e", width: "100%" },
  textarea: { fontSize: 13, fontFamily: "inherit", padding: "7px 10px", border: "1px solid #d1d5db", borderRadius: 7, background: "#fff", color: "#1a1a2e", width: "100%", resize: "vertical", minHeight: 70 },
  expEntry: { border: "1px solid #e5e7ef", borderRadius: 8, padding: 10, marginBottom: 8 },
  skillTag: { padding: "3px 9px", borderRadius: 20, fontSize: 11, background: "#f0f4ff", border: "1px solid #c7d2fe", display: "inline-flex", alignItems: "center", gap: 4 },
  kwChip: (added) => ({ padding: "3px 9px", borderRadius: 20, fontSize: 11, border: "1px solid " + (added ? "#2563eb" : "#d1d5db"), background: added ? "#eff6ff" : "#f9fafb", color: added ? "#2563eb" : "#374151", cursor: "pointer" }),
  stepTab: (active, done) => ({ padding: "5px 10px", fontSize: 11, border: "1px solid " + (active ? "#2563eb" : done ? "#16a34a" : "#d1d5db"), borderRadius: 7, cursor: "pointer", whiteSpace: "nowrap", background: active ? "#2563eb" : "transparent", color: active ? "#fff" : done ? "#16a34a" : "#6b7280", fontFamily: "inherit", transition: "all .15s" }),
  infoTip: { fontSize: 11, color: "#4b5563", padding: "6px 10px", background: "#eff6ff", borderRadius: 7, borderLeft: "3px solid #2563eb", marginBottom: 8 },
  sectionToggle: { display: "flex", alignItems: "center", gap: 8, padding: "6px 0", borderBottom: "1px solid #f3f4f6" },
  profCard: (active) => ({ padding: "8px 10px", border: "1px solid " + (active ? "#2563eb" : "#e5e7ef"), borderRadius: 8, cursor: "pointer", fontSize: 12, display: "flex", alignItems: "center", gap: 6, background: active ? "#eff6ff" : "#fff", color: active ? "#2563eb" : "#374151", fontWeight: active ? 600 : 400, transition: "all .15s" }),
  templateCard: (active) => ({ border: active ? "2px solid #2563eb" : "1px solid #e5e7ef", borderRadius: 8, padding: 8, cursor: "pointer", textAlign: "center", background: active ? "#eff6ff" : "#fff", transition: "all .15s" }),
  verItem: (active) => ({ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "7px 10px", border: "1px solid " + (active ? "#2563eb" : "#e5e7ef"), borderRadius: 8, fontSize: 12, cursor: "pointer", background: active ? "#eff6ff" : "#fff" }),
  jdKw: (match) => ({ padding: "2px 9px", borderRadius: 12, fontSize: 11, background: match ? "#dcfce7" : "#fef2f2", color: match ? "#15803d" : "#dc2626", border: "1px solid " + (match ? "#86efac" : "#fca5a5"), cursor: match ? "default" : "pointer" }),
  checkItem: (pass) => ({ fontSize: 11, display: "flex", alignItems: "center", gap: 5, color: pass ? "#16a34a" : "#dc2626" }),
};

// ─── RESUME PREVIEW RENDERERS ─────────────────────────────────────────────────

function ResumePreviewClassic({ data }) {
  const allSkills = [...data.skills.technical, ...data.skills.soft, ...data.skills.tools];
  return (
    <div style={{ fontFamily: "'Times New Roman', Times, serif", fontSize: 10, color: "#111", lineHeight: 1.5 }}>
      <div style={{ fontSize: 20, fontWeight: 700, marginBottom: 2 }}>{data.name || "Your Name"}</div>
      <div style={{ fontSize: 9, color: "#444", borderBottom: "1px solid #888", paddingBottom: 5, marginBottom: 8 }}>
        {[
          data.email,
          data.phone,
          data.location,
          data.linkedin ? "LinkedIn: " + (data.linkedin.match(/linkedin\.com\/in\/([^/]+)/i)?.[1] || data.linkedin).replace(/-?\d+$/, "").replace(/-/g, " ").replace(/\b\w/g, c => c.toUpperCase()) : "",
          data.website ? "GitHub: " + (data.website.match(/github\.com\/([^/]+)/i)?.[1] || data.website) : ""
        ].filter(Boolean).join("  |  ")}
      </div>
      {data.summary && <><div style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".5px", borderBottom: "1px solid #ccc", paddingBottom: 2, marginBottom: 4 }}>Professional Summary</div><p style={{ fontSize: 9, marginBottom: 8 }}>{data.summary}</p></>}
      {data.experience.some(e => e.title) && <>
        <div style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".5px", borderBottom: "1px solid #ccc", paddingBottom: 2, marginBottom: 4 }}>Work Experience</div>
        {data.experience.filter(e => e.title).map((e, i) => (
          <div key={i} style={{ marginBottom: 6 }}>
            <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 700, fontSize: 10 }}>
              <span>{e.title}{e.company ? ", " + e.company : ""}</span>
              <span style={{ fontWeight: 400 }}>{[e.startDate, e.endDate].filter(Boolean).join(" – ")}</span>
            </div>
            {e.location && <div style={{ fontSize: 9, color: "#555", marginBottom: 2 }}>{e.location}</div>}
            {e.bullets.filter(b => b).map((b, j) => <div key={j} style={{ fontSize: 9, paddingLeft: 10, position: "relative" }}><span style={{ position: "absolute", left: 0 }}>–</span>{b.split("\n").filter(l=>l.trim()).map((l,li)=><div key={li} style={{paddingLeft:10,fontSize:9,marginBottom:1,position:"relative"}}><span style={{position:"absolute",left:0}}>–</span>{l.replace(/^[-•]\s*/,"")}</div>)}</div>)}
          </div>
        ))}
      </>}
      {data.education.some(e => e.degree) && <>
        <div style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".5px", borderBottom: "1px solid #ccc", paddingBottom: 2, marginBottom: 4, marginTop: 8 }}>Education</div>
        {data.education.filter(e => e.degree).map((e, i) => (
          <div key={i} style={{ marginBottom: 4 }}>
            <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 700, fontSize: 10 }}><span>{e.degree}</span><span style={{ fontWeight: 400 }}>{e.year}</span></div>
            <div style={{ fontSize: 9, color: "#555" }}>{[e.institution, e.gpa ? (/b\.?tech|b\.?e\b/i.test(e.degree) ? "CGPA: " : "Percentage: ") + e.gpa : "", e.honors].filter(Boolean).join("  |  ")}</div>
          </div>
        ))}
      </>}
      {data.certifications && <><div style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".5px", borderBottom: "1px solid #ccc", paddingBottom: 2, marginBottom: 4, marginTop: 8 }}>Certifications</div>{data.certifications.split("\n").filter(Boolean).map((c, i) => <div key={i} style={{ fontSize: 9, paddingLeft: 10, position: "relative" }}><span style={{ position: "absolute", left: 0 }}>–</span>{c}</div>)}</>}
      {allSkills.length > 0 && <><div style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".5px", borderBottom: "1px solid #ccc", paddingBottom: 2, marginBottom: 4, marginTop: 8 }}>Skills</div><div style={{ fontSize: 9, display: "flex", flexWrap: "wrap", gap: 3 }}>{allSkills.map((s, i) => <span key={i} style={{ background: "#f0f0f0", padding: "1px 5px", borderRadius: 3 }}>{s}</span>)}</div></>}
    </div>
  );
}

function ResumePreviewModern({ data }) {
  const allSkills = [...data.skills.technical, ...data.skills.soft, ...data.skills.tools];
  return (
    <div style={{ fontFamily: "Arial, Helvetica, sans-serif", fontSize: 10, color: "#111", lineHeight: 1.5 }}>
      <div style={{ fontSize: 20, fontWeight: 700, textTransform: "uppercase", letterSpacing: 1, color: "#1e40af", marginBottom: 2 }}>{data.name || "YOUR NAME"}</div>
      <div style={{ fontSize: 9, color: "#555", marginBottom: 12 }}>{[
          data.email,
          data.phone,
          data.location,
          data.linkedin ? "LinkedIn: " + (data.linkedin.match(/linkedin\.com\/in\/([^/]+)/i)?.[1] || data.linkedin).replace(/-?\d+$/, "").replace(/-/g, " ").replace(/\b\w/g, c => c.toUpperCase()) : "",
          data.website ? "GitHub: " + (data.website.match(/github\.com\/([^/]+)/i)?.[1] || data.website) : ""
        ].filter(Boolean).join("  |  ")}</div>
      {data.summary && <><div style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".5px", color: "#1e40af", borderLeft: "3px solid #1e40af", paddingLeft: 6, marginBottom: 4 }}>Professional Summary</div><p style={{ fontSize: 9, marginBottom: 10 }}>{data.summary}</p></>}
      {data.experience.some(e => e.title) && <>
        <div style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase", color: "#1e40af", borderLeft: "3px solid #1e40af", paddingLeft: 6, marginBottom: 4 }}>Work Experience</div>
        {data.experience.filter(e => e.title).map((e, i) => (
          <div key={i} style={{ marginBottom: 6 }}>
            <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 700, fontSize: 10 }}><span>{e.title}{e.company ? ", " + e.company : ""}</span><span style={{ fontWeight: 400, color: "#555" }}>{[e.startDate, e.endDate].filter(Boolean).join(" – ")}</span></div>
            {e.bullets.filter(b => b).map((b, j) => <div key={j} style={{ fontSize: 9, paddingLeft: 10, position: "relative" }}><span style={{ position: "absolute", left: 0 }}>–</span>{b.split("\n").filter(l=>l.trim()).map((l,li)=><div key={li} style={{paddingLeft:10,fontSize:9,marginBottom:1,position:"relative"}}><span style={{position:"absolute",left:0}}>–</span>{l.replace(/^[-•]\s*/,"")}</div>)}</div>)}
          </div>
        ))}
      </>}
      {allSkills.length > 0 && <><div style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase", color: "#1e40af", borderLeft: "3px solid #1e40af", paddingLeft: 6, marginBottom: 4, marginTop: 8 }}>Skills</div><div style={{ fontSize: 9, display: "flex", flexWrap: "wrap", gap: 3 }}>{allSkills.map((s, i) => <span key={i} style={{ background: "#eff6ff", color: "#1e40af", padding: "1px 5px", borderRadius: 3 }}>{s}</span>)}</div></>}
      {data.education.some(e => e.degree) && <>
        <div style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase", color: "#1e40af", borderLeft: "3px solid #1e40af", paddingLeft: 6, marginBottom: 4, marginTop: 8 }}>Education</div>
        {data.education.filter(e => e.degree).map((e, i) => (
          <div key={i} style={{ marginBottom: 4 }}>
            <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 700, fontSize: 10 }}><span>{e.degree}</span><span style={{ fontWeight: 400, color: "#555" }}>{e.year}</span></div>
            <div style={{ fontSize: 9, color: "#555" }}>{[e.institution, e.university, e.gpa ? (/b\.?tech|b\.?e\b/i.test(e.degree) ? "CGPA: " : "Percentage: ") + e.gpa : "", /mbbs/i.test(e.degree) && e.mbbsClass ? "Class: "+e.mbbsClass : ""].filter(Boolean).join("  |  ")}</div>
          </div>
        ))}
      </>}
    </div>
  );
}

function ResumePreviewCreative({ data }) {
  const allSkills = [...data.skills.technical, ...data.skills.soft, ...data.skills.tools];
  return (
    <div style={{ fontFamily: "Arial, Helvetica, sans-serif", fontSize: 10, color: "#111", lineHeight: 1.5 }}>
      <div style={{ background: "#1e3a5f", color: "#fff", padding: "10px 14px", margin: "-20px -20px 12px", textAlign: "center" }}>
        <div style={{ fontSize: 18, fontWeight: 700 }}>{data.name || "Your Name"}</div>
        <div style={{ fontSize: 9, color: "rgba(255,255,255,.8)", marginTop: 3 }}>{[
          data.email,
          data.phone,
          data.location,
          data.linkedin ? "LinkedIn: " + (data.linkedin.match(/linkedin\.com\/in\/([^/]+)/i)?.[1] || data.linkedin).replace(/-?\d+$/, "").replace(/-/g, " ").replace(/\b\w/g, c => c.toUpperCase()) : "",
          data.website ? "GitHub: " + (data.website.match(/github\.com\/([^/]+)/i)?.[1] || data.website) : ""
        ].filter(Boolean).join("  |  ")}</div>
      </div>
      {data.summary && <><div style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase", color: "#1e3a5f", borderBottom: "1.5px solid #1e3a5f", paddingBottom: 2, marginBottom: 4 }}>Summary</div><p style={{ fontSize: 9, marginBottom: 8 }}>{data.summary}</p></>}
      {data.experience.some(e => e.title) && <>
        <div style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase", color: "#1e3a5f", borderBottom: "1.5px solid #1e3a5f", paddingBottom: 2, marginBottom: 4 }}>Experience</div>
        {data.experience.filter(e => e.title).map((e, i) => (
          <div key={i} style={{ marginBottom: 6 }}>
            <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 700, fontSize: 10 }}><span>{e.title}{e.company ? ", " + e.company : ""}</span><span style={{ fontWeight: 400, color: "#555" }}>{[e.startDate, e.endDate].filter(Boolean).join(" – ")}</span></div>
            {e.bullets.filter(b => b).map((b, j) => <div key={j} style={{ fontSize: 9, paddingLeft: 10, position: "relative" }}><span style={{ position: "absolute", left: 0 }}>–</span>{b.split("\n").filter(l=>l.trim()).map((l,li)=><div key={li} style={{paddingLeft:10,fontSize:9,marginBottom:1,position:"relative"}}><span style={{position:"absolute",left:0}}>–</span>{l.replace(/^[-•]\s*/,"")}</div>)}</div>)}
          </div>
        ))}
      </>}
      {allSkills.length > 0 && <><div style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase", color: "#1e3a5f", borderBottom: "1.5px solid #1e3a5f", paddingBottom: 2, marginBottom: 4, marginTop: 8 }}>Skills</div><div style={{ fontSize: 9, display: "flex", flexWrap: "wrap", gap: 3 }}>{allSkills.map((s, i) => <span key={i} style={{ background: "#dbeafe", color: "#1e40af", padding: "1px 5px", borderRadius: 3 }}>{s}</span>)}</div></>}
      {data.education.some(e => e.degree) && <>
        <div style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase", color: "#1e3a5f", borderBottom: "1.5px solid #1e3a5f", paddingBottom: 2, marginBottom: 4, marginTop: 8 }}>Education</div>
        {data.education.filter(e => e.degree).map((e, i) => (
          <div key={i} style={{ marginBottom: 4 }}>
            <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 700, fontSize: 10 }}><span>{e.degree}</span><span style={{ fontWeight: 400, color: "#555" }}>{e.year}</span></div>
            <div style={{ fontSize: 9, color: "#555" }}>{e.institution}</div>
          </div>
        ))}
      </>}
    </div>
  );
}

// ─── SCORE RING ───────────────────────────────────────────────────────────────

function ScoreRing({ score }) {
  const circ = 226;
  const offset = circ - (circ * score / 100);
  const color = score <= 40 ? "#dc2626" : score <= 70 ? "#d97706" : score <= 90 ? "#16a34a" : "#0d9488";
  const label = score <= 40 ? "Needs Work" : score <= 70 ? "Getting There" : score <= 90 ? "ATS Ready" : "Optimized";
  return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", padding: "8px 0" }}>
      <svg width="90" height="90" viewBox="0 0 90 90" aria-label={`ATS score: ${score}`}>
        <circle cx="45" cy="45" r="36" fill="none" stroke="#e5e7ef" strokeWidth="8" />
        <circle cx="45" cy="45" r="36" fill="none" stroke={color} strokeWidth="8" strokeLinecap="round"
          strokeDasharray={circ} strokeDashoffset={offset} transform="rotate(-90 45 45)"
          style={{ transition: "stroke-dashoffset .5s, stroke .5s" }} />
        <text x="45" y="50" textAnchor="middle" fontSize="18" fontWeight="700" fill="#1a1a2e">{score}</text>
      </svg>
      <div style={{ fontSize: 12, fontWeight: 600, color, marginTop: 2 }}>{label}</div>
    </div>
  );
}

// ─── TOGGLE ───────────────────────────────────────────────────────────────────

function Toggle({ on, onToggle, label }) {
  return (
    <div style={S.sectionToggle}>
      <div onClick={onToggle} role="switch" aria-checked={on} aria-label={label}
        style={{ width: 32, height: 18, background: on ? "#2563eb" : "#d1d5db", borderRadius: 9, position: "relative", cursor: "pointer", transition: "background .2s", flexShrink: 0 }}>
        <div style={{ width: 14, height: 14, background: "#fff", borderRadius: "50%", position: "absolute", top: 2, left: on ? 16 : 2, transition: "left .2s", boxShadow: "0 1px 2px rgba(0,0,0,.2)" }} />
      </div>
      <span style={{ fontSize: 13 }}>{label}</span>
    </div>
  );
}

// ─── MAIN COMPONENT ───────────────────────────────────────────────────────────

export default function ATSResumeGenerator() {
  // ── AUTOSAVE: Load from localStorage on first render ──────────────
  const loadSaved = () => {
    try {
      const saved = localStorage.getItem("ats_resume_data");
      return saved ? JSON.parse(saved) : emptyData();
    } catch { return emptyData(); }
  };
  const loadSavedMeta = (key, fallback) => {
    try {
      const saved = localStorage.getItem(key);
      return saved ? JSON.parse(saved) : fallback;
    } catch { return fallback; }
  };

  const [mode, setMode] = useState("A");
  const [prof, setProf] = useState(() => loadSavedMeta("ats_prof", "tech"));
  const [template, setTemplate] = useState(() => loadSavedMeta("ats_template", "classic"));
  const [step, setStep] = useState("personal");
  const [completedSteps, setCompletedSteps] = useState(new Set());
  const [data, setData] = useState(loadSaved);
  const [versions, setVersions] = useState(() => loadSavedMeta("ats_versions", []));
  const [pasteText, setPasteText] = useState("");
  const [parsing, setParsing] = useState(false);
  const [jdText, setJdText] = useState("");
  const [jdResults, setJdResults] = useState(null);
  const [wordExporting, setWordExporting] = useState(false);
  const [saveStatus, setSaveStatus] = useState("saved"); // "saved" | "saving" | "unsaved"
  const saveTimer = useRef(null);

  // ── AUTOSAVE: Save to localStorage whenever data changes ───────────
  useEffect(() => {
    setSaveStatus("unsaved");
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => {
      try {
        localStorage.setItem("ats_resume_data", JSON.stringify(data));
        localStorage.setItem("ats_prof", JSON.stringify(prof));
        localStorage.setItem("ats_template", JSON.stringify(template));
        localStorage.setItem("ats_versions", JSON.stringify(versions));
        setSaveStatus("saved");
      } catch { setSaveStatus("unsaved"); }
    }, 1000); // debounce 1 second
    return () => { if (saveTimer.current) clearTimeout(saveTimer.current); };
  }, [data, prof, template, versions]);

  const upd = useCallback((key, val) => setData(d => ({ ...d, [key]: val })), []);

  const updExp = useCallback((id, key, val) => setData(d => ({
    ...d, experience: d.experience.map(e => e.id === id ? { ...e, [key]: val } : e)
  })), []);

  const updBullet = useCallback((id, bi, val) => setData(d => ({
    ...d, experience: d.experience.map(e => e.id === id ? { ...e, bullets: e.bullets.map((b, i) => i === bi ? val : b) } : e)
  })), []);

  const updEdu = useCallback((id, key, val) => setData(d => ({
    ...d, education: d.education.map(e => e.id === id ? { ...e, [key]: val } : e)
  })), []);

  const addExp = () => setData(d => ({ ...d, experience: [...d.experience, { id: Date.now(), title: "", company: "", location: "", startDate: "", endDate: "", bullets: [""] }] }));
  const removeExp = (id) => setData(d => ({ ...d, experience: d.experience.filter(e => e.id !== id) }));
  const addBullet = (id) => setData(d => ({ ...d, experience: d.experience.map(e => e.id === id ? { ...e, bullets: [...e.bullets, ""] } : e) }));
  const removeBullet = (id, bi) => setData(d => ({ ...d, experience: d.experience.map(e => e.id === id ? { ...e, bullets: e.bullets.filter((_, i) => i !== bi) } : e) }));
  const addEdu = () => setData(d => ({ ...d, education: [...d.education, { id: Date.now(), degree: "", institution: "", year: "", gpa: "", honors: "", university: "", mbbsClass: "" }] }));
  const removeEdu = (id) => setData(d => ({ ...d, education: d.education.filter(e => e.id !== id) }));

  const addSkill = (cat) => {
    const val = (data.skillInput[cat] || "").trim();
    if (val && !data.skills[cat].includes(val)) {
      setData(d => ({ ...d, skills: { ...d.skills, [cat]: [...d.skills[cat], val] }, skillInput: { ...d.skillInput, [cat]: "" } }));
    }
  };
  const removeSkill = (cat, skill) => setData(d => ({ ...d, skills: { ...d.skills, [cat]: d.skills[cat].filter(s => s !== skill) } }));
  const addKWSkill = (kw) => {
    const all = [...data.skills.technical, ...data.skills.soft, ...data.skills.tools];
    if (!all.includes(kw)) setData(d => ({ ...d, skills: { ...d.skills, technical: [...d.skills.technical, kw] } }));
  };

  const goStep = (id) => { setCompletedSteps(s => new Set([...s, step])); setStep(id); };

  const toggleOptional = (id) => setData(d => ({ ...d, optionals: { ...d.optionals, [id]: !d.optionals[id] } }));

  const clearAllData = () => {
    if (window.confirm("Clear all saved data and start fresh?")) {
      localStorage.removeItem("ats_resume_data");
      localStorage.removeItem("ats_prof");
      localStorage.removeItem("ats_template");
      localStorage.removeItem("ats_versions");
      setData(emptyData());
      setProf("tech");
      setTemplate("classic");
      setVersions([]);
      setStep("personal");
      setSaveStatus("saved");
    }
  };

  const saveVersion = () => {
    const name = window.prompt("Version name (e.g. 'Google - SWE'):");
    if (!name) return;
    if (versions.length >= 5) { alert("Max 5 versions. Delete one first."); return; }
    setVersions(v => [...v, { id: Date.now(), name, data: JSON.parse(JSON.stringify(data)) }]);
  };
  const loadVersion = (id) => { const v = versions.find(v => v.id === id); if (v) setData(JSON.parse(JSON.stringify(v.data))); };
  const deleteVersion = (id) => setVersions(v => v.filter(x => x.id !== id));

  const parseResume = async () => {
    if (!pasteText.trim()) { alert("Paste resume text first."); return; }
    setParsing(true);
    try {
      const resp = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": import.meta.env.VITE_ANTHROPIC_API_KEY,
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true",
        },
        body: JSON.stringify({
          model: "claude-sonnet-4-5",
          max_tokens: 1000,
          system: 'You are a resume parser. Extract info and return ONLY valid JSON (no markdown, no backticks) with: {"name":"","email":"","phone":"","location":"","linkedin":"","website":"","summary":"","experience":[{"title":"","company":"","location":"","startDate":"","endDate":"","bullets":[""]}],"education":[{"degree":"","institution":"","year":"","gpa":"","honors":"","university":"","mbbsClass":""}],"skills":{"technical":[],"soft":[],"tools":[]},"certifications":""}',
          messages: [{ role: "user", content: "Parse this resume:\n\n" + pasteText }],
        }),
      });
      const json = await resp.json();
      const raw = json.content?.[0]?.text || "{}";
      const parsed = JSON.parse(raw.replace(/```json|```/g, "").trim());
      setData(d => ({
        ...d,
        name: parsed.name || "", email: parsed.email || "", phone: parsed.phone || "",
        location: parsed.location || "", linkedin: parsed.linkedin || "", website: parsed.website || "",
        summary: parsed.summary || "", certifications: parsed.certifications || "",
        experience: (parsed.experience || []).map((e, i) => ({ id: i + 1, ...e, bullets: e.bullets || [""] })),
        education: (parsed.education || []).map((e, i) => ({ id: i + 1, university: "", mbbsClass: "", ...e })),
        skills: { technical: parsed.skills?.technical || [], soft: parsed.skills?.soft || [], tools: parsed.skills?.tools || [] },
      }));
      setMode("A");
      setStep("personal");
      alert("Parsed! Review each section in Form Builder.");
    } catch (e) {
      alert("Parse failed: " + e.message);
    } finally {
      setParsing(false);
    }
  };

  const analyzeJD = () => {
    if (!jdText.trim()) return;
    const words = jdText.toLowerCase().match(/\b[\w#\+\.]+\b/g) || [];
    const freq = {};
    words.forEach(w => { if (w.length > 3) freq[w] = (freq[w] || 0) + 1; });
    const top = Object.entries(freq).sort((a, b) => b[1] - a[1]).slice(0, 20).map(e => e[0]);
    const allProfKws = Object.values(KW_PRESETS).flat().map(k => k.toLowerCase());
    const candidates = top.filter(w => allProfKws.some(k => k.includes(w) || w.includes(k.split(" ")[0]))).slice(0, 15);
    const resumeText = (data.summary + " " + [...data.skills.technical, ...data.skills.soft, ...data.skills.tools].join(" ") + " " + data.experience.flatMap(e => e.bullets).join(" ")).toLowerCase();
    const matched = candidates.filter(k => resumeText.includes(k));
    const missing = candidates.filter(k => !resumeText.includes(k));
    setJdResults({ matched, missing, pct: candidates.length > 0 ? Math.round(matched.length / candidates.length * 100) : 0 });
  };

  const addMissingKW = (kw) => {
    const proper = kw.charAt(0).toUpperCase() + kw.slice(1);
    if (!data.skills.technical.includes(proper)) setData(d => ({ ...d, skills: { ...d.skills, technical: [...d.skills.technical, proper] } }));
    if (jdText) setTimeout(analyzeJD, 100);
  };

  const handleWordExport = async () => {
    setWordExporting(true);
    try { await exportToWord(data); }
    catch (e) { alert("Word export failed: " + e.message + "\n\nMake sure 'docx' and 'file-saver' are installed."); }
    finally { setWordExporting(false); }
  };

  // Extract short display name from LinkedIn URL
  // e.g. https://www.linkedin.com/in/agnel-biju-280807290/ → agnel-biju-280807290
  const getLinkedInName = (url) => {
    if (!url) return "";
    const match = url.match(/linkedin\.com\/in\/([^/]+)/i);
    if (!match) return url;
    // Remove trailing digits and dashes, then capitalize each word
    const cleaned = match[1]
      .replace(/-?\d+$/, "")        // remove trailing digits like -280807290
      .replace(/-/g, " ")            // replace dashes with spaces
      .replace(/\b\w/g, c => c.toUpperCase()); // capitalize each word
    return cleaned;
  };

  // Extract GitHub username from URL
  // e.g. https://github.com/wizmedia_28 → github.com/wizmedia_28
  const getGitHubDisplay = (url) => {
    if (!url) return "";
    const match = url.match(/github\.com\/([^/]+)/i);
    return match ? match[1] : url;  // show only username
  };

  const printResume = () => {
    const allSkills = [...data.skills.technical, ...data.skills.soft, ...data.skills.tools];
    const getEduLabel = (degree) => /b\.?tech|b\.?e/i.test(degree) ? "CGPA" : "Percentage";
    
    // Build resume HTML based on selected template
    let resumeHTML = "";
    const tmpl = template;

    if (tmpl === "modern") {
      resumeHTML = `
        <div style="font-family:Arial,Helvetica,sans-serif;font-size:11pt;color:#111;line-height:1.5">
          <div style="font-size:22pt;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#1e40af;margin-bottom:4px">${data.name || ""}</div>
          <div style="font-size:10pt;color:#555;margin-bottom:16px;border-bottom:1px solid #ccc;padding-bottom:8px">${[
  data.email ? `<a href="mailto:${data.email}" style="color:inherit;text-decoration:none">${data.email}</a>` : "",
  data.phone ? `<a href="tel:${data.phone}" style="color:inherit;text-decoration:none">${data.phone}</a>` : "",
  data.location || "",
  data.linkedin ? `<a href="${data.linkedin.startsWith("http")?data.linkedin:"https://"+data.linkedin}" target="_blank" style="color:#0077b5;text-decoration:none"><svg xmlns=\"http://www.w3.org/2000/svg\" width=\"11\" height=\"11\" viewBox=\"0 0 24 24\" style=\"vertical-align:middle;margin-right:3px\"><path fill=\"#0077b5\" d=\"M20.447 20.452h-3.554v-5.569c0-1.328-.027-3.037-1.852-3.037-1.853 0-2.136 1.445-2.136 2.939v5.667H9.351V9h3.414v1.561h.046c.477-.9 1.637-1.85 3.37-1.85 3.601 0 4.267 2.37 4.267 5.455v6.286zM5.337 7.433a2.062 2.062 0 0 1-2.063-2.065 2.064 2.064 0 1 1 2.063 2.065zm1.782 13.019H3.555V9h3.564v11.452zM22.225 0H1.771C.792 0 0 .774 0 1.729v20.542C0 23.227.792 24 1.771 24h20.451C23.2 24 24 23.227 24 22.271V1.729C24 .774 23.2 0 22.222 0h.003z\"/></svg><span style=\"font-weight:700;font-size:9pt\">LinkedIn:</span> ${getLinkedInName(data.linkedin)}</a>` : "",
  data.website ? `<a href="${data.website.startsWith("http")?data.website:"https://"+data.website}" target="_blank" style="color:#333;text-decoration:none"><svg xmlns=\"http://www.w3.org/2000/svg\" width=\"11\" height=\"11\" viewBox=\"0 0 24 24\" style=\"vertical-align:middle;margin-right:3px\"><path fill=\"#333\" d=\"M12 0C5.374 0 0 5.373 0 12c0 5.302 3.438 9.8 8.207 11.387.599.111.793-.261.793-.577v-2.234c-3.338.726-4.033-1.416-4.033-1.416-.546-1.387-1.333-1.756-1.333-1.756-1.089-.745.083-.729.083-.729 1.205.084 1.839 1.237 1.839 1.237 1.07 1.834 2.807 1.304 3.492.997.107-.775.418-1.305.762-1.604-2.665-.305-5.467-1.334-5.467-5.931 0-1.311.469-2.381 1.236-3.221-.124-.303-.535-1.524.117-3.176 0 0 1.008-.322 3.301 1.23A11.509 11.509 0 0 1 12 5.803c1.02.005 2.047.138 3.006.404 2.291-1.552 3.297-1.23 3.297-1.23.653 1.653.242 2.874.118 3.176.77.84 1.235 1.911 1.235 3.221 0 4.609-2.807 5.624-5.479 5.929.43.372.823 1.102.823 2.222v3.293c0 .319.192.694.801.576C20.566 21.797 24 17.3 24 12c0-6.627-5.373-12-12-12z\"/></svg><span style=\"font-weight:700;font-size:9pt\">GitHub:</span> ${getGitHubDisplay(data.website)}</a>` : ""
].filter(Boolean).join("  |  ")}</div>
          ${data.summary ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;color:#1e40af;border-left:3px solid #1e40af;padding-left:8px;margin-bottom:6px">Professional Summary</div><p style="font-size:10pt;margin-bottom:14px">${data.summary}</p>` : ""}
          ${data.experience.some(e=>e.title) ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;color:#1e40af;border-left:3px solid #1e40af;padding-left:8px;margin-bottom:6px">Work Experience</div>${data.experience.filter(e=>e.title).map(e=>`<div style="margin-bottom:10px"><div style="display:flex;justify-content:space-between;font-weight:700;font-size:10pt"><span>${e.title}${e.company?", "+e.company:""}</span><span style="font-weight:400;color:#555">${[e.startDate,e.endDate].filter(Boolean).join(" – ")}</span></div>${e.bullets.filter(b=>b).map(b=>`${b.split("\n").filter(l=>l.trim()).map(l=>`<div style="padding-left:12px;font-size:9.5pt;margin-bottom:2px;position:relative"><span style="position:absolute;left:0">–</span>${l.replace(/^[-•]\s*/,"")}</div>`).join("")}`).join("")}</div>`).join("")}` : ""}
          ${(data.skills.technical.length||data.skills.soft.length||data.skills.tools.length) ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;color:#1e40af;border-left:3px solid #1e40af;padding-left:8px;margin-bottom:8px;margin-top:14px">Skills</div>
${data.skills.technical.length?`<div style="margin-bottom:5px"><span style="font-size:9.5pt;font-weight:700;color:#1e40af">Technical: </span><span style="font-size:9.5pt">${data.skills.technical.join(", ")}</span></div>`:""}
${data.skills.tools.length?`<div style="margin-bottom:5px"><span style="font-size:9.5pt;font-weight:700;color:#1e40af">Tools & Software: </span><span style="font-size:9.5pt">${data.skills.tools.join(", ")}</span></div>`:""}
${data.skills.soft.length?`<div style="margin-bottom:5px"><span style="font-size:9.5pt;font-weight:700;color:#1e40af">Soft Skills: </span><span style="font-size:9.5pt">${data.skills.soft.join(", ")}</span></div>`:""}` : ""}
          ${data.education.some(e=>e.degree) ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;color:#1e40af;border-left:3px solid #1e40af;padding-left:8px;margin-bottom:6px;margin-top:14px">Education</div>${data.education.filter(e=>e.degree).map(e=>`<div style="margin-bottom:6px"><div style="display:flex;justify-content:space-between;font-weight:700;font-size:10pt"><span>${e.degree}</span><span style="font-weight:400;color:#555">${e.year}</span></div><div style="font-size:9.5pt;color:#555">${[e.institution, e.university||"", e.gpa?getEduLabel(e.degree)+": "+e.gpa:"", /mbbs/i.test(e.degree)&&e.mbbsClass?"Class: "+e.mbbsClass:"", e.honors||""].filter(Boolean).join("  |  ")}</div></div>`).join("")}` : ""}
          ${data.optionals.projects&&Array.isArray(data.projects)&&data.projects.filter(Boolean).length ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;color:#1e40af;border-left:3px solid #1e40af;padding-left:8px;margin-bottom:6px;margin-top:14px">Projects</div>${data.projects.filter(Boolean).map(p=>{const[nm,ds,url]=p.split("|").map(s=>s.trim());return`<div style="margin-bottom:6px"><div style="font-weight:700;font-size:10pt">${nm||""}</div>${ds?`<div style="font-size:9.5pt;color:#555">${ds}${url?" | "+url:""}</div>`:""}</div>`}).join("")}` : ""}
          ${data.optionals.awards&&data.awards ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;color:#1e40af;border-left:3px solid #1e40af;padding-left:8px;margin-bottom:6px;margin-top:14px">Awards</div>${data.awards.split("\n").filter(Boolean).map(a=>`<div style="padding-left:12px;font-size:9.5pt;margin-bottom:2px;position:relative"><span style="position:absolute;left:0">–</span>${a}</div>`).join("")}` : ""}
          ${data.optionals.languages&&data.languages ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;color:#1e40af;border-left:3px solid #1e40af;padding-left:8px;margin-bottom:6px;margin-top:14px">Languages</div>${data.languages.split("\n").filter(Boolean).map(l=>`<div style="padding-left:12px;font-size:9.5pt;margin-bottom:2px;position:relative"><span style="position:absolute;left:0">–</span>${l}</div>`).join("")}` : ""}
        </div>`;
    } else if (tmpl === "creative") {
      resumeHTML = `
        <div style="font-family:Arial,Helvetica,sans-serif;font-size:11pt;color:#111;line-height:1.5">
          <div style="background:#1e3a5f;color:#fff;padding:20px 28px;margin:-28px -28px 18px;text-align:center">
            <div style="font-size:22pt;font-weight:700">${data.name || ""}</div>
            <div style="font-size:10pt;color:rgba(255,255,255,.85);margin-top:5px">${[
  data.email ? `<a href="mailto:${data.email}" style="color:inherit;text-decoration:none">${data.email}</a>` : "",
  data.phone ? `<a href="tel:${data.phone}" style="color:inherit;text-decoration:none">${data.phone}</a>` : "",
  data.location || "",
  data.linkedin ? `<a href="${data.linkedin.startsWith("http")?data.linkedin:"https://"+data.linkedin}" target="_blank" style="color:#0077b5;text-decoration:none"><svg xmlns=\"http://www.w3.org/2000/svg\" width=\"11\" height=\"11\" viewBox=\"0 0 24 24\" style=\"vertical-align:middle;margin-right:3px\"><path fill=\"#0077b5\" d=\"M20.447 20.452h-3.554v-5.569c0-1.328-.027-3.037-1.852-3.037-1.853 0-2.136 1.445-2.136 2.939v5.667H9.351V9h3.414v1.561h.046c.477-.9 1.637-1.85 3.37-1.85 3.601 0 4.267 2.37 4.267 5.455v6.286zM5.337 7.433a2.062 2.062 0 0 1-2.063-2.065 2.064 2.064 0 1 1 2.063 2.065zm1.782 13.019H3.555V9h3.564v11.452zM22.225 0H1.771C.792 0 0 .774 0 1.729v20.542C0 23.227.792 24 1.771 24h20.451C23.2 24 24 23.227 24 22.271V1.729C24 .774 23.2 0 22.222 0h.003z\"/></svg><span style=\"font-weight:700;font-size:9pt\">LinkedIn:</span> ${getLinkedInName(data.linkedin)}</a>` : "",
  data.website ? `<a href="${data.website.startsWith("http")?data.website:"https://"+data.website}" target="_blank" style="color:#333;text-decoration:none"><svg xmlns=\"http://www.w3.org/2000/svg\" width=\"11\" height=\"11\" viewBox=\"0 0 24 24\" style=\"vertical-align:middle;margin-right:3px\"><path fill=\"#333\" d=\"M12 0C5.374 0 0 5.373 0 12c0 5.302 3.438 9.8 8.207 11.387.599.111.793-.261.793-.577v-2.234c-3.338.726-4.033-1.416-4.033-1.416-.546-1.387-1.333-1.756-1.333-1.756-1.089-.745.083-.729.083-.729 1.205.084 1.839 1.237 1.839 1.237 1.07 1.834 2.807 1.304 3.492.997.107-.775.418-1.305.762-1.604-2.665-.305-5.467-1.334-5.467-5.931 0-1.311.469-2.381 1.236-3.221-.124-.303-.535-1.524.117-3.176 0 0 1.008-.322 3.301 1.23A11.509 11.509 0 0 1 12 5.803c1.02.005 2.047.138 3.006.404 2.291-1.552 3.297-1.23 3.297-1.23.653 1.653.242 2.874.118 3.176.77.84 1.235 1.911 1.235 3.221 0 4.609-2.807 5.624-5.479 5.929.43.372.823 1.102.823 2.222v3.293c0 .319.192.694.801.576C20.566 21.797 24 17.3 24 12c0-6.627-5.373-12-12-12z\"/></svg><span style=\"font-weight:700;font-size:9pt\">GitHub:</span> ${getGitHubDisplay(data.website)}</a>` : ""
].filter(Boolean).join("  |  ")}</div>
          </div>
          ${data.summary ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;color:#1e3a5f;border-bottom:2px solid #1e3a5f;padding-bottom:3px;margin-bottom:7px">Summary</div><p style="font-size:10pt;margin-bottom:14px">${data.summary}</p>` : ""}
          ${data.experience.some(e=>e.title) ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;color:#1e3a5f;border-bottom:2px solid #1e3a5f;padding-bottom:3px;margin-bottom:7px">Experience</div>${data.experience.filter(e=>e.title).map(e=>`<div style="margin-bottom:10px"><div style="display:flex;justify-content:space-between;font-weight:700;font-size:10pt"><span>${e.title}${e.company?", "+e.company:""}</span><span style="font-weight:400;color:#555">${[e.startDate,e.endDate].filter(Boolean).join(" – ")}</span></div>${e.bullets.filter(b=>b).map(b=>`${b.split("\n").filter(l=>l.trim()).map(l=>`<div style="padding-left:12px;font-size:9.5pt;margin-bottom:2px;position:relative"><span style="position:absolute;left:0">–</span>${l.replace(/^[-•]\s*/,"")}</div>`).join("")}`).join("")}</div>`).join("")}` : ""}
          ${(data.skills.technical.length||data.skills.soft.length||data.skills.tools.length) ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;color:#1e3a5f;border-bottom:2px solid #1e3a5f;padding-bottom:3px;margin-bottom:8px;margin-top:14px">Skills</div>
${data.skills.technical.length?`<div style="margin-bottom:5px"><span style="font-size:9.5pt;font-weight:700;color:#1e3a5f">Technical: </span><span style="font-size:9.5pt">${data.skills.technical.join(", ")}</span></div>`:""}
${data.skills.tools.length?`<div style="margin-bottom:5px"><span style="font-size:9.5pt;font-weight:700;color:#1e3a5f">Tools & Software: </span><span style="font-size:9.5pt">${data.skills.tools.join(", ")}</span></div>`:""}
${data.skills.soft.length?`<div style="margin-bottom:5px"><span style="font-size:9.5pt;font-weight:700;color:#1e3a5f">Soft Skills: </span><span style="font-size:9.5pt">${data.skills.soft.join(", ")}</span></div>`:""}` : ""}
          ${data.education.some(e=>e.degree) ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;color:#1e3a5f;border-bottom:2px solid #1e3a5f;padding-bottom:3px;margin-bottom:7px;margin-top:14px">Education</div>${data.education.filter(e=>e.degree).map(e=>`<div style="margin-bottom:6px"><div style="display:flex;justify-content:space-between;font-weight:700;font-size:10pt"><span>${e.degree}</span><span style="font-weight:400;color:#555">${e.year}</span></div><div style="font-size:9.5pt;color:#555">${[e.institution, e.university||"", e.gpa?getEduLabel(e.degree)+": "+e.gpa:"", /mbbs/i.test(e.degree)&&e.mbbsClass?"Class: "+e.mbbsClass:"", e.honors||""].filter(Boolean).join("  |  ")}</div></div>`).join("")}` : ""}
          ${data.optionals.projects&&Array.isArray(data.projects)&&data.projects.filter(Boolean).length ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;color:#1e3a5f;border-bottom:2px solid #1e3a5f;padding-bottom:3px;margin-bottom:7px;margin-top:14px">Projects</div>${data.projects.filter(Boolean).map(p=>{const[nm,ds,url]=p.split("|").map(s=>s.trim());return`<div style="margin-bottom:6px"><div style="font-weight:700;font-size:10pt">${nm||""}</div>${ds?`<div style="font-size:9.5pt;color:#555">${ds}${url?" | "+url:""}</div>`:""}</div>`}).join("")}` : ""}
          ${data.optionals.awards&&data.awards ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;color:#1e3a5f;border-bottom:2px solid #1e3a5f;padding-bottom:3px;margin-bottom:7px;margin-top:14px">Awards</div>${data.awards.split("\n").filter(Boolean).map(a=>`<div style="padding-left:12px;font-size:9.5pt;margin-bottom:2px;position:relative"><span style="position:absolute;left:0">–</span>${a}</div>`).join("")}` : ""}
          ${data.optionals.languages&&data.languages ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;color:#1e3a5f;border-bottom:2px solid #1e3a5f;padding-bottom:3px;margin-bottom:7px;margin-top:14px">Languages</div>${data.languages.split("\n").filter(Boolean).map(l=>`<div style="padding-left:12px;font-size:9.5pt;margin-bottom:2px;position:relative"><span style="position:absolute;left:0">–</span>${l}</div>`).join("")}` : ""}
        </div>`;
    } else {
      // Classic and Executive
      const secColor = tmpl === "executive" ? "#1e3a5f" : "#111";
      resumeHTML = `
        <div style="font-family:'Times New Roman',Times,serif;font-size:11pt;color:#111;line-height:1.5">
          <div style="font-size:22pt;font-weight:700;margin-bottom:4px;text-align:center">${data.name || ""}</div>
          <div style="font-size:10pt;color:#444;border-bottom:1px solid #888;padding-bottom:8px;margin-bottom:12px;text-align:center">${[
  data.email ? `<a href="mailto:${data.email}" style="color:inherit;text-decoration:none">${data.email}</a>` : "",
  data.phone ? `<a href="tel:${data.phone}" style="color:inherit;text-decoration:none">${data.phone}</a>` : "",
  data.location || "",
  data.linkedin ? `<a href="${data.linkedin.startsWith("http")?data.linkedin:"https://"+data.linkedin}" target="_blank" style="color:#0077b5;text-decoration:none"><svg xmlns=\"http://www.w3.org/2000/svg\" width=\"11\" height=\"11\" viewBox=\"0 0 24 24\" style=\"vertical-align:middle;margin-right:3px\"><path fill=\"#0077b5\" d=\"M20.447 20.452h-3.554v-5.569c0-1.328-.027-3.037-1.852-3.037-1.853 0-2.136 1.445-2.136 2.939v5.667H9.351V9h3.414v1.561h.046c.477-.9 1.637-1.85 3.37-1.85 3.601 0 4.267 2.37 4.267 5.455v6.286zM5.337 7.433a2.062 2.062 0 0 1-2.063-2.065 2.064 2.064 0 1 1 2.063 2.065zm1.782 13.019H3.555V9h3.564v11.452zM22.225 0H1.771C.792 0 0 .774 0 1.729v20.542C0 23.227.792 24 1.771 24h20.451C23.2 24 24 23.227 24 22.271V1.729C24 .774 23.2 0 22.222 0h.003z\"/></svg><span style=\"font-weight:700;font-size:9pt\">LinkedIn:</span> ${getLinkedInName(data.linkedin)}</a>` : "",
  data.website ? `<a href="${data.website.startsWith("http")?data.website:"https://"+data.website}" target="_blank" style="color:#333;text-decoration:none"><svg xmlns=\"http://www.w3.org/2000/svg\" width=\"11\" height=\"11\" viewBox=\"0 0 24 24\" style=\"vertical-align:middle;margin-right:3px\"><path fill=\"#333\" d=\"M12 0C5.374 0 0 5.373 0 12c0 5.302 3.438 9.8 8.207 11.387.599.111.793-.261.793-.577v-2.234c-3.338.726-4.033-1.416-4.033-1.416-.546-1.387-1.333-1.756-1.333-1.756-1.089-.745.083-.729.083-.729 1.205.084 1.839 1.237 1.839 1.237 1.07 1.834 2.807 1.304 3.492.997.107-.775.418-1.305.762-1.604-2.665-.305-5.467-1.334-5.467-5.931 0-1.311.469-2.381 1.236-3.221-.124-.303-.535-1.524.117-3.176 0 0 1.008-.322 3.301 1.23A11.509 11.509 0 0 1 12 5.803c1.02.005 2.047.138 3.006.404 2.291-1.552 3.297-1.23 3.297-1.23.653 1.653.242 2.874.118 3.176.77.84 1.235 1.911 1.235 3.221 0 4.609-2.807 5.624-5.479 5.929.43.372.823 1.102.823 2.222v3.293c0 .319.192.694.801.576C20.566 21.797 24 17.3 24 12c0-6.627-5.373-12-12-12z\"/></svg><span style=\"font-weight:700;font-size:9pt\">GitHub:</span> ${getGitHubDisplay(data.website)}</a>` : ""
].filter(Boolean).join("  |  ")}</div>
          ${data.summary ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:${secColor};border-bottom:1px solid #ccc;padding-bottom:2px;margin-bottom:6px">Professional Summary</div><p style="font-size:10pt;margin-bottom:12px">${data.summary}</p>` : ""}
          ${data.experience.some(e=>e.title) ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:${secColor};border-bottom:1px solid #ccc;padding-bottom:2px;margin-bottom:6px">Work Experience</div>${data.experience.filter(e=>e.title).map(e=>`<div style="margin-bottom:10px"><div style="display:flex;justify-content:space-between;font-weight:700;font-size:10pt"><span>${e.title}${e.company?", "+e.company:""}</span><span style="font-weight:400">${[e.startDate,e.endDate].filter(Boolean).join(" – ")}</span></div>${e.location?`<div style="font-size:9.5pt;color:#555;margin-bottom:2px">${e.location}</div>`:""} ${e.bullets.filter(b=>b).map(b=>`${b.split("\n").filter(l=>l.trim()).map(l=>`<div style="padding-left:12px;font-size:9.5pt;margin-bottom:2px;position:relative"><span style="position:absolute;left:0">–</span>${l.replace(/^[-•]\s*/,"")}</div>`).join("")}`).join("")}</div>`).join("")}` : ""}
          ${data.education.some(e=>e.degree) ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:${secColor};border-bottom:1px solid #ccc;padding-bottom:2px;margin-bottom:6px;margin-top:12px">Education</div>${data.education.filter(e=>e.degree).map(e=>`<div style="margin-bottom:8px"><div style="display:flex;justify-content:space-between;font-weight:700;font-size:10pt"><span>${e.degree}</span><span style="font-weight:400">${e.year}</span></div><div style="font-size:9.5pt;color:#555">${[e.institution,e.gpa?getEduLabel(e.degree)+": "+e.gpa:"",e.honors].filter(Boolean).join("  |  ")}</div></div>`).join("")}` : ""}
          ${data.certifications ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:${secColor};border-bottom:1px solid #ccc;padding-bottom:2px;margin-bottom:6px;margin-top:12px">Certifications</div>${data.certifications.split("\n").filter(Boolean).map(c=>`<div style="padding-left:12px;font-size:9.5pt;margin-bottom:2px;position:relative"><span style="position:absolute;left:0">–</span>${c}</div>`).join("")}` : ""}
          ${(data.skills.technical.length||data.skills.soft.length||data.skills.tools.length) ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:${secColor};border-bottom:1px solid #ccc;padding-bottom:2px;margin-bottom:6px;margin-top:12px">Skills</div>
${data.skills.technical.length?`<div style="margin-bottom:4px"><span style="font-size:9.5pt;font-weight:700;color:#333">Technical: </span><span style="font-size:9.5pt">${data.skills.technical.join(", ")}</span></div>`:""}
${data.skills.tools.length?`<div style="margin-bottom:4px"><span style="font-size:9.5pt;font-weight:700;color:#333">Tools & Software: </span><span style="font-size:9.5pt">${data.skills.tools.join(", ")}</span></div>`:""}
${data.skills.soft.length?`<div style="margin-bottom:4px"><span style="font-size:9.5pt;font-weight:700;color:#333">Soft Skills: </span><span style="font-size:9.5pt">${data.skills.soft.join(", ")}</span></div>`:""}` : ""}
          ${data.optionals.projects&&Array.isArray(data.projects)&&data.projects.filter(Boolean).length ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:${secColor};border-bottom:1px solid #ccc;padding-bottom:2px;margin-bottom:6px;margin-top:12px">Projects</div>${data.projects.filter(Boolean).map(p=>{const[nm,ds,url]=p.split("|").map(s=>s.trim());return`<div style="margin-bottom:6px"><div style="font-weight:700;font-size:10pt">${nm||""}</div>${ds?`<div style="font-size:9.5pt;color:#555">${ds}${url?" | "+url:""}</div>`:""}</div>`}).join("")}` : ""}
          ${data.optionals.awards&&data.awards ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:${secColor};border-bottom:1px solid #ccc;padding-bottom:2px;margin-bottom:6px;margin-top:12px">Awards & Honors</div>${data.awards.split("\n").filter(Boolean).map(a=>`<div style="padding-left:12px;font-size:9.5pt;margin-bottom:2px;position:relative"><span style="position:absolute;left:0">–</span>${a}</div>`).join("")}` : ""}
          ${data.optionals.languages&&data.languages ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:${secColor};border-bottom:1px solid #ccc;padding-bottom:2px;margin-bottom:6px;margin-top:12px">Languages</div>${data.languages.split("\n").filter(Boolean).map(l=>`<div style="padding-left:12px;font-size:9.5pt;margin-bottom:2px;position:relative"><span style="position:absolute;left:0">–</span>${l}</div>`).join("")}` : ""}
          ${data.optionals.volunteer&&data.volunteer ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:${secColor};border-bottom:1px solid #ccc;padding-bottom:2px;margin-bottom:6px;margin-top:12px">Volunteer Work</div><p style="font-size:9.5pt">${data.volunteer}</p>` : ""}
          ${data.optionals.publications&&data.publications ? `<div style="font-size:10pt;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:${secColor};border-bottom:1px solid #ccc;padding-bottom:2px;margin-bottom:6px;margin-top:12px">Publications</div>${data.publications.split("\n").filter(Boolean).map(p=>`<div style="padding-left:12px;font-size:9.5pt;margin-bottom:2px;position:relative"><span style="position:absolute;left:0">–</span>${p}</div>`).join("")}` : ""}
        </div>`;
    }

    const printWindow = window.open("", "_blank", "width=900,height=700");
    printWindow.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><title>${data.name || "Resume"} - Resume</title>
      <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { padding: 25mm 20mm; background: white; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        @media print { body { padding: 20mm; } @page { margin: 0; size: A4; } }
      </style>
    </head><body>${resumeHTML}<script>window.onload=function(){window.print();}<\/script></body></html>`);
    printWindow.document.close();
  };

  const downloadHTML = () => {
    const allSkills = [...data.skills.technical, ...data.skills.soft, ...data.skills.tools];
    const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${data.name || "Resume"}</title>
<style>body{font-family:'Times New Roman',serif;max-width:800px;margin:0 auto;padding:40px;font-size:11pt;color:#111;line-height:1.5}h1{font-size:20pt;margin-bottom:4px}.contact{font-size:9pt;color:#444;border-bottom:1px solid #aaa;padding-bottom:8px;margin-bottom:12px}h2{font-size:11pt;text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid #ccc;padding-bottom:2px;margin:12px 0 6px}.entry{margin-bottom:8px}.entry-header{display:flex;justify-content:space-between;font-weight:700}.entry-sub{font-size:9pt;color:#555;margin-bottom:3px}.bullet{padding-left:12px;font-size:9pt;margin-bottom:2px;position:relative}.bullet::before{content:"–";position:absolute;left:0}.skills{font-size:9pt}@media print{body{padding:20mm}}</style>
</head><body>
<h1>${data.name || "Your Name"}</h1>
<div class="contact">${[data.email, data.phone, data.location, data.linkedin, data.website].filter(Boolean).join(" | ")}</div>
${data.summary ? `<h2>Professional Summary</h2><p>${data.summary}</p>` : ""}
${data.experience.some(e => e.title) ? `<h2>Work Experience</h2>${data.experience.filter(e => e.title).map(e => `<div class="entry"><div class="entry-header"><span>${e.title}${e.company ? ", " + e.company : ""}</span><span style="font-weight:400">${[e.startDate, e.endDate].filter(Boolean).join(" – ")}</span></div>${e.location ? `<div class="entry-sub">${e.location}</div>` : ""}${e.bullets.filter(b => b).map(b => `<div class="bullet">${b.split("\n").filter(l=>l.trim()).map((l,li)=><div key={li} style={{paddingLeft:10,fontSize:9,marginBottom:1,position:"relative"}}><span style={{position:"absolute",left:0}}>–</span>{l.replace(/^[-•]\s*/,"")}</div>)}</div>`).join("")}</div>`).join("")}` : ""}
${data.education.some(e => e.degree) ? `<h2>Education</h2>${data.education.filter(e => e.degree).map(e => `<div class="entry"><div class="entry-header"><span>${e.degree}</span><span style="font-weight:400">${e.year}</span></div><div class="entry-sub">${[e.institution, e.university||"", e.gpa?(/b\.?tech|b\.?e\b/i.test(e.degree)?"CGPA: ":"Percentage: ")+e.gpa:"", /mbbs/i.test(e.degree)&&e.mbbsClass?"Class: "+e.mbbsClass:""].filter(Boolean).join(" | ")}</div></div>`).join("")}` : ""}
${allSkills.length ? `<h2>Skills</h2><div class="skills">${allSkills.join(" | ")}</div>` : ""}
</body></html>`;
    const blob = new Blob([html], { type: "text/html" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = (data.name || "resume").replace(/\s+/g, "_") + "_resume.html";
    a.click();
  };

  const { score, checks } = calcScore(data, prof);
  const ResumePreview = template === "modern" ? ResumePreviewModern : template === "creative" ? ResumePreviewCreative : ResumePreviewClassic;

  // ─── STEP CONTENT ───────────────────────────────────────────────────────────

  const renderStep = () => {
    if (step === "personal") return (
      <div style={S.card}>
        <div style={S.cardTitle}>👤 Personal Info</div>
        <div style={S.formRow(2)} className="form-row-2">
          <div style={S.formGroup}><label style={S.label}>Full Name</label><input style={S.input} value={data.name} onChange={e => upd("name", e.target.value)} placeholder="Jane Smith" /></div>
          <div style={S.formGroup}><label style={S.label}>Email</label><input style={S.input} type="email" value={data.email} onChange={e => upd("email", e.target.value)} placeholder="jane@email.com" /></div>
        </div>
        <div style={S.formRow(2)} className="form-row-2">
          <div style={S.formGroup}><label style={S.label}>Phone</label><input style={S.input} value={data.phone} onChange={e => upd("phone", e.target.value)} placeholder="+1 (555) 000-0000" /></div>
          <div style={S.formGroup}><label style={S.label}>Location</label><input style={S.input} value={data.location} onChange={e => upd("location", e.target.value)} placeholder="City, State" /></div>
        </div>
        <div style={S.formRow(2)} className="form-row-2">
          <div style={S.formGroup}><label style={S.label}>LinkedIn</label><input style={S.input} value={data.linkedin} onChange={e => upd("linkedin", e.target.value)} placeholder="linkedin.com/in/jane" /></div>
          <div style={S.formGroup}><label style={S.label}>GitHub</label><input style={S.input} value={data.website} onChange={e => upd("website", e.target.value)} placeholder="github.com/username" /></div>
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 8 }}>
          <button style={{ ...S.btn, ...S.btnPrimary }} onClick={() => goStep("summary")}>Next: Summary →</button>
        </div>
      </div>
    );

    if (step === "summary") return (
      <div style={S.card}>
        <div style={S.cardTitle}>📝 Professional Summary</div>
        <div style={S.infoTip}>Aim for 40–80 words. Suggested verbs: {(ACTION_VERBS[prof] || []).slice(0, 6).join(", ")}.</div>
        <div style={S.formGroup}>
          <label style={S.label}>Summary ({data.summary.split(/\s+/).filter(w => w).length} words)</label>
          <textarea style={{ ...S.textarea, minHeight: 90 }} value={data.summary} onChange={e => upd("summary", e.target.value)} placeholder="Results-driven professional with X years of experience..." />
        </div>
        <div style={{ marginBottom: 8 }}>
          <div style={{ fontSize: 11, fontWeight: 600, color: "#6b7280", marginBottom: 5 }}>KEYWORD SUGGESTIONS</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
            {(KW_PRESETS[prof] || []).slice(0, 8).map(k => (
              <span key={k} style={S.kwChip(false)} onClick={() => upd("summary", (data.summary + " " + k).trim())}>{k}</span>
            ))}
          </div>
        </div>
        <div style={{ display: "flex", justifyContent: "space-between", marginTop: 8 }}>
          <button style={S.btn} onClick={() => goStep("personal")}>← Back</button>
          <button style={{ ...S.btn, ...S.btnPrimary }} onClick={() => goStep("experience")}>Next: Experience →</button>
        </div>
      </div>
    );

    if (step === "experience") return (
      <div style={S.card}>
        <div style={S.cardTitle}>🏢 Work Experience</div>
        <div style={S.infoTip}>Strong action verbs: {(ACTION_VERBS[prof] || []).join(", ")}</div>
        {data.experience.map((e, idx) => (
          <div key={e.id} style={S.expEntry}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <span style={{ fontSize: 12, fontWeight: 600 }}>{e.title || "Experience " + (idx + 1)}</span>
              {data.experience.length > 1 && <button style={{ ...S.btn, ...S.btnSm, borderColor: "#dc2626", color: "#dc2626" }} onClick={() => removeExp(e.id)}>🗑</button>}
            </div>
            <div style={S.formRow(2)} className="form-row-2">
              <div style={S.formGroup}><label style={S.label}>Job Title</label><input style={S.input} value={e.title} onChange={ev => updExp(e.id, "title", ev.target.value)} placeholder="Software Engineer" /></div>
              <div style={S.formGroup}><label style={S.label}>Company</label><input style={S.input} value={e.company} onChange={ev => updExp(e.id, "company", ev.target.value)} placeholder="Acme Corp" /></div>
            </div>
            <div style={S.formRow(3)} className="form-row-3">
              <div style={S.formGroup}><label style={S.label}>Location</label><input style={S.input} value={e.location} onChange={ev => updExp(e.id, "location", ev.target.value)} placeholder="City, ST" /></div>
              <div style={S.formGroup}><label style={S.label}>Start Date</label><input style={S.input} value={e.startDate} onChange={ev => updExp(e.id, "startDate", ev.target.value)} placeholder="Jan 2022" /></div>
              <div style={S.formGroup}><label style={S.label}>End Date</label><input style={S.input} value={e.endDate} onChange={ev => updExp(e.id, "endDate", ev.target.value)} placeholder="Present" /></div>
            </div>
            <label style={{ ...S.label, display: "block", marginBottom: 3 }}>Bullet Points</label>
            {e.bullets.map((b, bi) => (
              <div key={bi} style={{ display: "flex", gap: 6, alignItems: "flex-start", marginBottom: 4 }}>
                <textarea style={{ ...S.textarea, minHeight: 38, flex: 1 }} value={b} onChange={ev => updBullet(e.id, bi, ev.target.value)} placeholder="Describe your achievement with impact..." />
                {e.bullets.length > 1 && <button style={{ ...S.btn, ...S.btnSm }} onClick={() => removeBullet(e.id, bi)}>×</button>}
              </div>
            ))}
            <button style={{ ...S.btn, ...S.btnSm }} onClick={() => addBullet(e.id)}>+ Add Bullet</button>
          </div>
        ))}
        <button style={{ ...S.btn, width: "100%", justifyContent: "center" }} onClick={addExp}>+ Add Experience</button>
        <div style={{ display: "flex", justifyContent: "space-between", marginTop: 10 }}>
          <button style={S.btn} onClick={() => goStep("summary")}>← Back</button>
          <button style={{ ...S.btn, ...S.btnPrimary }} onClick={() => goStep("education")}>Next: Education →</button>
        </div>
      </div>
    );

    if (step === "education") return (
      <div style={S.card}>
        <div style={S.cardTitle}>🎓 Education & Certifications</div>
        {data.education.map((e, idx) => (
          <div key={e.id} style={S.expEntry}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <span style={{ fontSize: 12, fontWeight: 600 }}>{e.degree || "Education " + (idx + 1)}</span>
              {data.education.length > 1 && <button style={{ ...S.btn, ...S.btnSm, borderColor: "#dc2626", color: "#dc2626" }} onClick={() => removeEdu(e.id)}>🗑</button>}
            </div>
            <div style={S.formRow(2)} className="form-row-2">
              <div style={S.formGroup}><label style={S.label}>Degree</label><input style={S.input} value={e.degree} onChange={ev => updEdu(e.id, "degree", ev.target.value)} placeholder="B.Sc. Computer Science" /></div>
              <div style={S.formGroup}><label style={S.label}>Institution</label><input style={S.input} value={e.institution} onChange={ev => updEdu(e.id, "institution", ev.target.value)} placeholder="MIT" /></div>
            </div>
            <div style={S.formRow(3)} className="form-row-3">
              <div style={S.formGroup}><label style={S.label}>Year</label><input style={S.input} value={e.year} onChange={ev => updEdu(e.id, "year", ev.target.value)} placeholder="2020" /></div>
              <div style={S.formGroup}><label style={S.label}>{/b\.?tech|b\.?e\b/i.test(e.degree) ? "CGPA (optional)" : "Percentage (optional)"}</label><input style={S.input} value={e.gpa} onChange={ev => updEdu(e.id, "gpa", ev.target.value)} placeholder={/b\.?tech|b\.?e\b/i.test(e.degree) ? "8.5" : "85%"} /></div>
              <div style={S.formGroup}><label style={S.label}>Honors (optional)</label><input style={S.input} value={e.honors} onChange={ev => updEdu(e.id, "honors", ev.target.value)} placeholder="Cum Laude" /></div>
            </div>
            <div style={S.formRow(2)} className="form-row-2">
              <div style={S.formGroup}><label style={S.label}>University (optional)</label><input style={S.input} value={e.university||""} onChange={ev => updEdu(e.id, "university", ev.target.value)} placeholder="University of Kerala" /></div>
              {/mbbs/i.test(e.degree) && (
                <div style={S.formGroup}>
                  <label style={S.label}>Class (Distinction / First Class)</label>
                  <select style={S.input} value={e.mbbsClass||""} onChange={ev => updEdu(e.id, "mbbsClass", ev.target.value)}>
                    <option value="">Select Class</option>
                    <option value="Distinction">Distinction</option>
                    <option value="First Class">First Class</option>
                    <option value="Second Class">Second Class</option>
                    <option value="Pass">Pass</option>
                  </select>
                </div>
              )}
            </div>
            <div style={{display:"none"}}>
            </div>
          </div>
        ))}
        <button style={{ ...S.btn, width: "100%", justifyContent: "center" }} onClick={addEdu}>+ Add Education</button>
        <div style={{ ...S.formRow(1), marginTop: 10 }}>
          <div style={S.formGroup}><label style={S.label}>Certifications (one per line)</label><textarea style={S.textarea} value={data.certifications} onChange={e => upd("certifications", e.target.value)} placeholder={"AWS Certified Solutions Architect, 2023\nPMP Certification, 2022"} /></div>
        </div>
        <div style={{ display: "flex", justifyContent: "space-between", marginTop: 8 }}>
          <button style={S.btn} onClick={() => goStep("experience")}>← Back</button>
          <button style={{ ...S.btn, ...S.btnPrimary }} onClick={() => goStep("skills")}>Next: Skills →</button>
        </div>
      </div>
    );

    if (step === "skills") {
      const allAdded = [...data.skills.technical, ...data.skills.soft, ...data.skills.tools];
      return (
        <div style={S.card}>
          <div style={S.cardTitle}>🏷 Skills</div>
          <div style={S.infoTip}>Click chips to add profession keywords, or type your own.</div>
          <div style={{ marginBottom: 12 }}>
            <div style={{ fontSize: 11, fontWeight: 600, color: "#6b7280", marginBottom: 5 }}>SUGGESTED KEYWORDS</div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
              {(KW_PRESETS[prof] || []).map(k => <span key={k} style={S.kwChip(allAdded.includes(k))} onClick={() => addKWSkill(k)}>{k}</span>)}
            </div>
          </div>
          {[["technical", "Technical Skills"], ["soft", "Soft Skills"], ["tools", "Tools & Software"]].map(([cat, lbl]) => (
            <div key={cat} style={{ marginBottom: 12 }}>
              <label style={S.label}>{lbl}</label>
              <div style={{ display: "flex", gap: 6, marginTop: 3 }}>
                <input style={S.input} value={data.skillInput[cat]} onChange={e => setData(d => ({ ...d, skillInput: { ...d.skillInput, [cat]: e.target.value } }))} onKeyDown={e => { if (e.key === "Enter") { e.preventDefault(); addSkill(cat); } }} placeholder="Type and press Enter" />
                <button style={{ ...S.btn, ...S.btnSm }} onClick={() => addSkill(cat)}>+</button>
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 5, marginTop: 5 }}>
                {(data.skills[cat] || []).map(s => <span key={s} style={S.skillTag}>{s}<span style={{ cursor: "pointer", color: "#6b7280", marginLeft: 2 }} onClick={() => removeSkill(cat, s)}>×</span></span>)}
              </div>
            </div>
          ))}
          <div style={{ display: "flex", justifyContent: "space-between", marginTop: 8 }}>
            <button style={S.btn} onClick={() => goStep("education")}>← Back</button>
            <button style={{ ...S.btn, ...S.btnPrimary }} onClick={() => goStep("optional")}>Next: Optional →</button>
          </div>
        </div>
      );
    }

    if (step === "optional") return (
      <div style={S.card}>
        <div style={S.cardTitle}>📋 Optional Sections</div>
        {[["projects","Projects"],["publications","Publications"],["volunteer","Volunteer Work"],["awards","Awards & Honors"],["languages","Languages"]].map(([id, lbl]) => (
          <Toggle key={id} on={data.optionals[id]} onToggle={() => toggleOptional(id)} label={lbl} />
        ))}
        <div style={{ marginTop: 12 }}>
          {data.optionals.projects && <div style={S.formGroup}><label style={S.label}>Projects (Name | Description | URL)</label><textarea style={S.textarea} value={Array.isArray(data.projects) ? data.projects.join("\n") : ""} onChange={e => upd("projects", e.target.value.split("\n"))} placeholder={"Project Name | Brief description | github.com/..."} /></div>}
          {data.optionals.volunteer && <div style={{ ...S.formGroup, marginTop: 8 }}><label style={S.label}>Volunteer Work</label><textarea style={S.textarea} value={data.volunteer} onChange={e => upd("volunteer", e.target.value)} /></div>}
          {data.optionals.awards && <div style={{ ...S.formGroup, marginTop: 8 }}><label style={S.label}>Awards & Honors (one per line)</label><textarea style={S.textarea} value={data.awards} onChange={e => upd("awards", e.target.value)} /></div>}
          {data.optionals.languages && <div style={{ ...S.formGroup, marginTop: 8 }}><label style={S.label}>Languages (one per line)</label><textarea style={S.textarea} value={data.languages} onChange={e => upd("languages", e.target.value)} placeholder={"Spanish – Fluent\nFrench – Conversational"} /></div>}
          {data.optionals.publications && <div style={{ ...S.formGroup, marginTop: 8 }}><label style={S.label}>Publications (one per line)</label><textarea style={S.textarea} value={data.publications} onChange={e => upd("publications", e.target.value)} /></div>}
        </div>
        <div style={{ display: "flex", justifyContent: "flex-start", marginTop: 10 }}>
          <button style={S.btn} onClick={() => goStep("skills")}>← Back</button>
        </div>
      </div>
    );
  };

  // ─── RENDER ────────────────────────────────────────────────────────────────

  return (
    <div style={S.app}>
      {/* Topbar */}
      <div style={S.topbar}>
        <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={S.appTitle}>📄 ATS Resume Generator</div>
            <div style={{
              fontSize: 10, padding: "2px 8px", borderRadius: 10,
              background: saveStatus === "saved" ? "#dcfce7" : saveStatus === "saving" ? "#fef9c3" : "#fee2e2",
              color: saveStatus === "saved" ? "#16a34a" : saveStatus === "saving" ? "#ca8a04" : "#dc2626",
              fontWeight: 500, display: "flex", alignItems: "center", gap: 3
            }}>
              {saveStatus === "saved" ? "✓ Auto-saved" : saveStatus === "saving" ? "⏳ Saving..." : "● Unsaved"}
            </div>
          </div>
          <div style={S.modeTabs}>
            <button style={S.modeTab(mode === "A")} onClick={() => setMode("A")}>Form Builder</button>
            <button style={S.modeTab(mode === "B")} onClick={() => setMode("B")}>Paste & Parse</button>
          </div>
        </div>
        <div style={S.btnGroup}>
          <button style={{ ...S.btn, fontSize: 11, borderColor: "#dc2626", color: "#dc2626" }} onClick={clearAllData} title="Clear all data">
            🗑 Clear
          </button>
          <button style={{ ...S.btn, ...S.btnWord }} onClick={handleWordExport} disabled={wordExporting}>
            {wordExporting ? "Exporting..." : "⬇ .docx"}
          </button>
          <button style={S.btn} onClick={downloadHTML}>⬇ HTML</button>
          <button style={{ ...S.btn, ...S.btnPrimary }} onClick={() => printResume()}>🖨 PDF</button>
          <button id="mobile-preview-btn" style={{ ...S.btn, display: "none", background: "#7c3aed", color: "#fff", border: "1px solid #7c3aed" }}
            onClick={() => {
              const pc = document.getElementById("preview-col");
              const mb = document.getElementById("mobile-preview-btn");
              if (pc.style.display === "flex") {
                pc.style.display = "none";
                mb.textContent = "👁 Preview";
              } else {
                pc.style.display = "flex";
                pc.style.flexDirection = "column";
                mb.textContent = "✏️ Edit";
              }
            }}>👁 Preview</button>
        </div>
      </div>

      <div style={S.main} id="main-area">
        {/* Editor Column */}
        <div style={S.editorCol} id="editor-col">
          {mode === "B" && (
            <div style={S.card}>
              <div style={S.cardTitle}>📋 Paste Your Resume</div>
              <p style={{ fontSize: 12, color: "#6b7280", marginBottom: 8 }}>Paste your existing resume as plain text. The AI will extract all sections and auto-fill the form.</p>
              <div style={{ position: "relative" }}>
                <textarea style={{ ...S.textarea, minHeight: 160, fontSize: 12 }} value={pasteText} onChange={e => setPasteText(e.target.value)} placeholder="Paste resume text here..." />
                {parsing && (
                  <div style={{ position: "absolute", inset: 0, background: "rgba(255,255,255,.88)", display: "flex", alignItems: "center", justifyContent: "center", flexDirection: "column", gap: 8, borderRadius: 7, fontSize: 13, fontWeight: 500 }}>
                    <div style={{ width: 24, height: 24, border: "2px solid #d1d5db", borderTopColor: "#2563eb", borderRadius: "50%", animation: "spin .8s linear infinite" }} />
                    Parsing your resume...
                  </div>
                )}
              </div>
              <div style={{ display: "flex", gap: 8, marginTop: 8, alignItems: "center" }}>
                <button style={{ ...S.btn, ...S.btnPrimary }} onClick={parseResume} disabled={parsing}>✨ Parse & Auto-Fill</button>
                <span style={{ fontSize: 11, color: "#6b7280" }}>Then switch to Form Builder to review</span>
              </div>
            </div>
          )}

          {mode === "A" && (
            <>
              <div style={S.card}>
                <div style={S.cardTitle}>🧭 Profession</div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
                  {PROFESSIONS.map(p => (
                    <div key={p.id} style={S.profCard(prof === p.id)} onClick={() => setProf(p.id)}>
                      {p.icon} {p.label}
                    </div>
                  ))}
                </div>
              </div>

              <div style={S.card}>
                <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
                  {STEPS.map(s => (
                    <button key={s.id} style={S.stepTab(step === s.id, completedSteps.has(s.id) && step !== s.id)} onClick={() => goStep(s.id)}>{s.label}</button>
                  ))}
                </div>
              </div>

              {renderStep()}
            </>
          )}
        </div>

        {/* Preview Column */}
        <div style={S.previewCol} id="preview-col">
          {/* ATS Score */}
          <div style={S.card}>
            <div style={S.cardTitle}>📊 ATS Score</div>
            <ScoreRing score={score} />
            <div style={{ display: "flex", flexDirection: "column", gap: 4, marginTop: 8 }}>
              {checks.map((c, i) => (
                <div key={i} style={S.checkItem(c.pass)}>
                  <span>{c.pass ? "✓" : "✗"}</span> {c.text}
                </div>
              ))}
            </div>
          </div>

          {/* Template Selector */}
          <div style={S.card}>
            <div style={S.cardTitle}>🎨 Template</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
              {TEMPLATES.map(t => (
                <div key={t.id} style={S.templateCard(template === t.id)} onClick={() => setTemplate(t.id)}>
                  <div style={{ height: 44, background: t.id === "creative" ? "#1e3a5f" : t.id === "modern" ? "#eff6ff" : "#f9fafb", borderRadius: 4, marginBottom: 4, border: "1px solid #e5e7ef" }} />
                  <div style={{ fontSize: 11, fontWeight: 600 }}>{t.name}</div>
                  <div style={{ fontSize: 10, color: "#6b7280" }}>{t.sub}</div>
                </div>
              ))}
            </div>
          </div>

          {/* Versions */}
          <div style={S.card}>
            <div style={S.cardTitle}>📁 Versions <span style={{ fontSize: 10, color: "#6b7280" }}>(up to 5)</span></div>
            {versions.length === 0 && <p style={{ fontSize: 11, color: "#9ca3af", padding: "4px 0" }}>No saved versions yet.</p>}
            <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
              {versions.map(v => (
                <div key={v.id} style={S.verItem(false)}>
                  <span style={{ flex: 1, cursor: "pointer" }} onClick={() => loadVersion(v.id)}>{v.name}</span>
                  <button style={{ ...S.btn, ...S.btnSm, borderColor: "#dc2626", color: "#dc2626" }} onClick={() => deleteVersion(v.id)}>🗑</button>
                </div>
              ))}
            </div>
            <button style={{ ...S.btn, width: "100%", justifyContent: "center", marginTop: 6 }} onClick={saveVersion}>+ Save Current Version</button>
          </div>

          {/* JD Analyzer */}
          <div style={S.card}>
            <div style={S.cardTitle}>🔍 JD Keyword Analyzer</div>
            <p style={{ fontSize: 11, color: "#6b7280", marginBottom: 6 }}>Paste a job description to see which keywords you're missing.</p>
            <textarea style={{ ...S.textarea, minHeight: 70, fontSize: 12 }} value={jdText} onChange={e => setJdText(e.target.value)} placeholder="Paste job description here..." />
            <button style={{ ...S.btn, ...S.btnPrimary, ...S.btnSm, marginTop: 6 }} onClick={analyzeJD}>Analyze JD</button>
            {jdResults && (
              <div style={{ marginTop: 8 }}>
                <div style={{ fontSize: 22, fontWeight: 600, color: "#2563eb", textAlign: "center" }}>{jdResults.pct}%</div>
                <div style={{ fontSize: 11, color: "#6b7280", textAlign: "center", marginBottom: 8 }}>keyword match ({jdResults.matched.length}/{jdResults.matched.length + jdResults.missing.length} found)</div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
                  {jdResults.matched.map(k => <span key={k} style={S.jdKw(true)}>✓ {k}</span>)}
                  {jdResults.missing.map(k => <span key={k} style={S.jdKw(false)} onClick={() => addMissingKW(k)} title="Click to add">+ {k}</span>)}
                </div>
                {jdResults.missing.length > 0 && <p style={{ fontSize: 11, color: "#6b7280", marginTop: 6 }}>Click red keywords to add them to your skills.</p>}
              </div>
            )}
          </div>

          {/* Live Preview */}
          <div style={S.card}>
            <div style={S.cardTitle}>👁 Live Preview</div>
            <div style={{ transform: "scale(0.72)", transformOrigin: "top left", width: "138.9%", pointerEvents: "none" }}>
              <div id="print-resume" style={{ background: "#fff", padding: 20, boxShadow: "0 2px 12px rgba(0,0,0,.12)", minHeight: 500, borderRadius: 4 }}>
                <ResumePreview data={data} />
              </div>
            </div>
          </div>
        </div>
      </div>

      <style>{`
        @keyframes spin { to { transform: rotate(360deg) } }
        * { box-sizing: border-box; }

        /* ── MOBILE ── */
        @media (max-width: 768px) {
          #main-area { flex-direction: column !important; }
          #preview-col { display: none !important; }
          #editor-col { padding: 10px !important; }
          #mobile-preview-btn { display: flex !important; }
          .topbar-buttons span { display: none; }
          .prof-grid { grid-template-columns: 1fr 1fr !important; }
          .form-row-2 { grid-template-columns: 1fr !important; }
          .form-row-3 { grid-template-columns: 1fr 1fr !important; }
          .step-tabs { flex-wrap: wrap !important; gap: 4px !important; }
        }

        @media (max-width: 480px) {
          .form-row-3 { grid-template-columns: 1fr !important; }
          .topbar-title { font-size: 13px !important; }
        }

        @media print {
          * { visibility: hidden; }
          #print-resume, #print-resume * { visibility: visible; }
          #print-resume {
            position: fixed;
            top: 0; left: 0;
            width: 100%; height: 100%;
            padding: 20mm;
            box-sizing: border-box;
            background: white;
            transform: none !important;
            font-size: 11pt;
          }
        }
      `}</style>
    </div>
  );
}
