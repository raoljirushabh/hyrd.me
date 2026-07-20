import { useState, useRef, useEffect } from "react";
import * as mammoth from "mammoth";

const C = {
  bg: "#0d1117",
  card: "#161b22",
  border: "#21262d",
  muted: "#6e7681",
  accent: "#6c63ff",
  text: "#e6edf3",
  green: "#3fb950",
  yellow: "#d29922",
  red: "#f85149",
};

export default function App() {
  const [mode, setMode] = useState("paste");
  const [jd, setJd] = useState("");
  const [resText, setResText] = useState("");
  const [fileTxt, setFileTxt] = useState("");
  const [fileName, setFileName] = useState("");
  const [questions, setQuestions] = useState([""]);
  const [loading, setLoading] = useState(false);
  const [loadMsg, setLoadMsg] = useState("");
  const [result, setResult] = useState(null);
  const [tab, setTab] = useState("r");
  const [error, setError] = useState("");
  const [copied, setCopied] = useState({ r: false, c: false, q: false });
  const [atsAnim, setAtsAnim] = useState(false);
  const [drag, setDrag] = useState(false);
  const ticker = useRef(null);

  useEffect(() => {
    const load = (src) =>
      new Promise((res, rej) => {
        if (document.querySelector(`script[src="${src}"]`)) return res();
        const s = document.createElement("script");
        s.src = src;
        s.onload = res;
        s.onerror = rej;
        document.head.appendChild(s);
      });
    load(
      "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"
    ).then(() => {
      window.pdfjsLib &&
        (window.pdfjsLib.GlobalWorkerOptions.workerSrc =
          "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js");
    });
    load("https://cdnjs.cloudflare.com/ajax/libs/docx/7.8.2/docx.umd.min.js");
  }, []);

  useEffect(() => {
    if (result) setTimeout(() => setAtsAnim(true), 120);
    else setAtsAnim(false);
  }, [result]);

  const MSGS = [
    "Analyzing job description...",
    "Extracting key requirements...",
    "Tailoring your resume...",
    "Weaving in the right keywords...",
    "Crafting your cover letter...",
    "Answering application questions...",
    "Scoring ATS compatibility...",
  ];

  const handleFile = async (file) => {
    if (!file) return;
    setFileName(file.name);
    setFileTxt("");
    const ab = await file.arrayBuffer();
    try {
      if (file.name.toLowerCase().endsWith(".docx")) {
        const r = await mammoth.extractRawText({ arrayBuffer: ab });
        setFileTxt(r.value);
      } else {
        const pdf = await window.pdfjsLib.getDocument({ data: ab }).promise;
        let t = "";
        for (let i = 1; i <= pdf.numPages; i++) {
          const pg = await pdf.getPage(i);
          const ct = await pg.getTextContent();
          t += ct.items.map((x) => x.str).join(" ") + "\n";
        }
        setFileTxt(t);
      }
    } catch {
      setError("Could not parse file. Try pasting the text instead.");
    }
  };

  const sanitize = (txt) => {
    if (!txt) return txt;
    return (
      txt
        // Em dashes and double hyphens
        .replace(/\s*—\s*/g, ", ")
        .replace(/\s*--\s*/g, ", ")
        // Semicolons: replace with period and capitalise next word
        .replace(/;\s*([a-zA-Z])/g, (_, c) => ". " + c.toUpperCase())
        // Mid-sentence colons (skip lines that look like resume section headers or time values)
        .replace(
          /([a-z]{3,}[^:\n]{2,})\s*:\s*([a-zA-Z])/g,
          (match, before, after) => {
            // Leave time patterns e.g. 12:30 and ratio patterns e.g. 3:1
            return before + ". " + after.toUpperCase();
          }
        )
        // Contractions — full forms only
        .replace(/\bI've\b/g, "I have")
        .replace(/\bI'd\b/g, "I would")
        .replace(/\bI'm\b/g, "I am")
        .replace(/\bI'll\b/g, "I will")
        .replace(/\bdon't\b/gi, "do not")
        .replace(/\bcan't\b/gi, "cannot")
        .replace(/\bwon't\b/gi, "will not")
        .replace(/\bdidn't\b/gi, "did not")
        .replace(/\bdoesn't\b/gi, "does not")
        .replace(/\bwouldn't\b/gi, "would not")
        .replace(/\bcouldn't\b/gi, "could not")
        .replace(/\bshouldn't\b/gi, "should not")
        .replace(/\bwasn't\b/gi, "was not")
        .replace(/\bweren't\b/gi, "were not")
        .replace(/\bhasn't\b/gi, "has not")
        .replace(/\bhaven't\b/gi, "have not")
        .replace(/\bhadn't\b/gi, "had not")
        .replace(/\bisn't\b/gi, "is not")
        .replace(/\baren't\b/gi, "are not")
        .replace(/\bit's\b/gi, "it is")
        .replace(/\bthat's\b/gi, "that is")
        .replace(/\bthere's\b/gi, "there is")
        .replace(/\bwhat's\b/gi, "what is")
        .replace(/\bwho's\b/gi, "who is")
        .replace(/\bhe's\b/gi, "he is")
        .replace(/\bshe's\b/gi, "she is")
        .replace(/\bthey're\b/gi, "they are")
        .replace(/\bwe're\b/gi, "we are")
        .replace(/\byou're\b/gi, "you are")
        .replace(/\bthey've\b/gi, "they have")
        .replace(/\bwe've\b/gi, "we have")
        .replace(/\byou've\b/gi, "you have")
        .replace(/\bhe'd\b/gi, "he would")
        .replace(/\bshe'd\b/gi, "she would")
        .replace(/\bthey'd\b/gi, "they would")
        .replace(/\bwe'd\b/gi, "we would")
        .replace(/\byou'd\b/gi, "you would")
        .replace(/\bhe'll\b/gi, "he will")
        .replace(/\bshe'll\b/gi, "she will")
        .replace(/\bthey'll\b/gi, "they will")
        .replace(/\bwe'll\b/gi, "we will")
        .replace(/\byou'll\b/gi, "you will")
        .replace(/\blet's\b/gi, "let us")
    );
  };

  const addQ = () => setQuestions((q) => [...q, ""]);
  const removeQ = (i) => setQuestions((q) => q.filter((_, j) => j !== i));
  const updateQ = (i, v) =>
    setQuestions((q) => q.map((x, j) => (j === i ? v : x)));
  const activeQs = questions.filter((q) => q.trim());

  const generate = async () => {
    setError("");
    const resume = mode === "paste" ? resText.trim() : fileTxt.trim();
    if (!jd.trim()) {
      setError("Please paste a job description.");
      return;
    }
    if (!resume) {
      setError("Please provide your resume.");
      return;
    }
    setLoading(true);
    setResult(null);
    let i = 0;
    setLoadMsg(MSGS[0]);
    ticker.current = setInterval(() => {
      i = (i + 1) % MSGS.length;
      setLoadMsg(MSGS[i]);
    }, 2200);

    const qSection =
      activeQs.length > 0
        ? `\n\nAPPLICATION QUESTIONS TO ANSWER:\n${activeQs
            .map((q, i) => `${i + 1}. ${q}`)
            .join("\n")}`
        : "";

    const prompt = `You are a world-class career strategist and ghostwriter. You have spent 20 years helping senior professionals, executives, and career switchers land roles at companies they genuinely want to work at. Your writing is indistinguishable from a sharp, self-aware human who knows how to tell their story without sounding like a brochure.

═══ MISSION ═══
Produce a resume, cover letter${
      activeQs.length > 0 ? ", and application question answers" : ""
    } that together score as close to 100% ATS match as possible — while reading like they were written by the candidate themselves on a good day. Not an AI. Not a career coach template. A real person.

═══ HUMAN-LIKE AUTHENTICITY — NON-NEGOTIABLE ═══
COVER LETTER:
- The opening line must be genuinely unexpected. Not "I am writing to express my interest." Not "I have long admired your company." Open with a sharp observation about the company, the industry challenge, or the role itself that shows you have actually thought about this. One or two sentences, then pivot to the candidate.
- Use contractions naturally where they would appear in confident professional writing: "I have", "I would", "it is", "there is". Not forced — just enough to break the corporate glaze.
- Vary sentence length dramatically. Short punchy sentence. Then one that runs a little longer, adds context, and shows the thinking behind a decision or result the candidate is proud of.
- Transitions must be invisible. Never: "Furthermore", "Moreover", "Additionally", "In conclusion", "To summarize", "It is worth noting", "I am confident that", "I believe that", "Needless to say", "With that said".
- Include one specific detail that shows the candidate knows this company — not flattery, but evidence of genuine attention.
- Close with something direct and human: "Happy to dig into this on a call" or "I would welcome the chance to talk through how I could contribute from day one." Not "I look forward to hearing from you at your earliest convenience."

RESUME BULLETS:
- Every bullet should read like a person describing their best work to someone they respect — precise, clear, a little proud without being boastful.
- No two bullets in the same role should start with the same verb or follow the same grammatical structure. Variety is a human signal.
- Mix bullet lengths: some punchy and short (one strong result), some slightly longer (context + action + outcome). Never all the same.
- Banned phrases: "Utilized", "Facilitated", "Interfaced with", "Spearheaded initiatives", "Drove synergistic outcomes", "Leveraged", "Proactively", "Robust", "Best-in-class", "Value-add". Use plain, direct, strong verbs instead.
- No bullet should start with "I".

APPLICATION QUESTION ANSWERS:
- Answer as the candidate in a confident, thoughtful interview — first person, grounded, specific to both the JD and their background.
- 2-4 sentences per answer. Direct. Not a rehearsed speech. Not an essay. Just a real, considered answer.
- Behavioral questions get a concise story with an outcome. Motivation questions get genuine reasoning — one specific thing about the company or role, not generic enthusiasm. Skills questions get a concrete example.
- Never open with "Great question", "I am passionate about", "I have always been driven by", "Throughout my career". Just start with the answer.

═══ GRAMMAR, TYPO & CONSISTENCY (apply everywhere) ═══
- Fix all grammatical errors silently: subject-verb agreement, tense (past for past roles, present for current), articles, prepositions.
- Standardize dates: "Mon YYYY – Mon YYYY" with en dash (–). "Present" capitalized for current roles.
- Bullet punctuation: consistent throughout — either all end with a period or none do. Prefer no period for fragments.
- Capitalize technologies and tools correctly: JavaScript, AWS, Python, TypeScript, etc.
- Remove redundancies: "in order to" → "to", "due to the fact that" → "because", "was able to" → just the verb.
- Numbers: spell out one–nine in prose; numerals for 10+, all percentages, dollar amounts, measurements.

═══ ATS MAXIMUM COVERAGE ═══
SUMMARY (4-5 lines): Mirror the JD — use the exact job title from the posting, top 4-5 required skills, candidate is most relevant experience framed around the role is core need.

SKILLS SECTION: List every hard skill, tool, technology, methodology, and domain keyword in the JD that the candidate can reasonably claim. This is the highest-leverage ATS section.

KEYWORD AUDIT — STRATEGIC PLACEMENT:
Every JD keyword must appear somewhere in the resume — but placement is everything. Follow this hierarchy:
- TIER 1 (highest-impact) keywords — the job title itself, the top 3-4 required skills, and any keyword that appears more than twice in the JD — must appear in the Professional Summary AND in at least one bullet in the most relevant role. These are what ATS systems weight most heavily.
- TIER 2 keywords — tools, methodologies, domain terms, soft skills mentioned in requirements — belong in the Skills section and woven into 1-2 bullets across roles where they fit naturally.
- TIER 3 keywords — keywords from the "nice to have" or company culture sections — seed lightly into the summary or a single bullet. Do not force them.
- Placement subtlety rules: never drop a keyword at the start of a bullet (it reads as planted). Bury it mid-sentence where it flows. Never cluster more than 2 keywords in the same bullet. If a keyword appears awkward, restructure the sentence until it disappears into the meaning. The test: read it aloud — if the keyword sticks out, rewrite.
- Mirror JD phrasing exactly where it sounds natural ("cross-functional stakeholder management" not "managing stakeholders across functions"). ATS matches strings.
- Acronyms: introduce with full form + acronym on first use, then acronym only after (e.g., "Infrastructure as Code (IaC)"). This captures both forms in ATS indexing.

STAR FORMAT BULLETS — SUBTLE JD INTEGRATION:
For every key responsibility or requirement in the JD that is not already covered by an existing bullet, add a new bullet in STAR format — but written so the structure is invisible. The goal is a bullet that reads like a natural recollection of real work, not a template.

How to construct a STAR bullet covertly:
- Situation/Task: one or two words of implicit context embedded mid-sentence, never a separate clause ("Facing rapid user growth..." or "Tasked with reducing onboarding friction...")
- Action: the main verb phrase — strong, specific, uses the JD is Tier 1 or Tier 2 keyword naturally embedded here
- Result: a concrete outcome, quantified if the resume provides any basis, otherwise scoped ("across a 12-person team", "within a 6-week sprint", "reducing escalations noticeably")
- Full bullet reads as one fluid sentence or two short ones — never three clauses joined by semicolons, which is a template tell

STAR bullet examples of what this looks like when done right:
✓ "Rebuilt the client onboarding pipeline after persistent drop-off issues, cutting time-to-activation by 30% and lifting 90-day retention across the SMB segment."
✓ "Took on ownership of vendor contract renewals during a team restructure, renegotiating terms with three key partners and recovering $180K in annual spend."
✗ "Situation: Faced with high churn. Task: Improve retention. Action: Implemented new process. Result: Reduced churn by 20%." ← never do this

Ground every added bullet in the candidate is actual role context. The Situation/Task must be plausible for their title, industry, and seniority. Never invent a specific metric, project name, or tool not present in the original resume — scope and scale instead ("a team of eight", "enterprise-scale deployment") where quantification is not possible.

Target: 5-7 bullets for the most recent/relevant role, 4-5 for mid-history roles, 2-3 for older ones.

═══ HARD RULES ═══
1. NO em dashes (—), double hyphens (--), or ellipses used for effect (...). Restructure the sentence.
2. NO fabrication — enhance and reframe only. Never invent metrics, credentials, or projects not in the original.
3. NO AI vocabulary (full list above). If a word appears on the banned list, delete it and find a plain human alternative.
4. NO uniform sentence or bullet structure throughout. Variation is authenticity.
5. Return ONLY valid JSON. No markdown, no fences, no commentary before or after.

FINAL SELF-CHECK BEFORE OUTPUT — scan every sentence in your draft and verify:
- Zero em dashes (—) or double hyphens (--)
- Zero semicolons (;) used in any sentence
- Zero colons (:) used mid-sentence to introduce a list or clause
- Zero contractions anywhere — "I have" not "I have", "I would" not "I would", "it is" not "it is", "do not" not "do not", "cannot" not "cannot", "will not" not "will not", and so on
- Every bullet in the same role starts with a different verb
- No two consecutive sentences follow identical structure
If any of the above are present, rewrite the offending sentence before returning.

JOB DESCRIPTION:
${jd}

CANDIDATE'S RESUME:
${resume}${qSection}

Return ONLY this exact JSON (nothing before or after):
{
  "tailored_resume": "full rewritten resume as plain text",
  "cover_letter": "full cover letter as plain text, 3-4 focused paragraphs",
  "application_answers": ${
    activeQs.length > 0
      ? `[${activeQs
          .map((q) => `{"question":"${q.replace(/"/g, '\\"')}","answer":"..."}`)
          .join(",")}]`
      : "[]"
  },
  "ats_score": 95,
  "ats_confidence": "Strong",
  "ats_summary": "One precise sentence — if below 95, state exactly what is still missing",
  "ats_breakdown": { "Keyword Match": 95, "Skills Alignment": 93, "Experience Relevance": 90, "Formatting & Clarity": 96 },
  "keywords_added": ["up to 10 keywords naturally woven in from the JD"]
}`;

    try {
      const resp = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true",
        },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 6000,
          messages: [{ role: "user", content: prompt }],
        }),
      });
      const data = await resp.json();
      if (data.error) throw new Error(data.error.message);
      const raw = data.content[0].text
        .replace(/^```json\s*/, "")
        .replace(/```\s*$/, "")
        .trim();
      const parsed = JSON.parse(raw);
      const cleaned = {
        ...parsed,
        tailored_resume: sanitize(parsed.tailored_resume),
        cover_letter: sanitize(parsed.cover_letter),
        application_answers: (parsed.application_answers || []).map((qa) => ({
          ...qa,
          answer: sanitize(qa.answer),
        })),
      };
      setResult(cleaned);
      setTab("r");
    } catch (e) {
      setError("Something went wrong: " + e.message);
    } finally {
      clearInterval(ticker.current);
      setLoading(false);
    }
  };

  const cp = async (t) => {
    let txt = "";
    if (t === "r") txt = result.tailored_resume;
    else if (t === "c") txt = result.cover_letter;
    else
      txt = (result.application_answers || [])
        .map((a) => `Q: ${a.question}\n\nA: ${a.answer}`)
        .join("\n\n---\n\n");
    await navigator.clipboard.writeText(txt).catch(() => {});
    setCopied((p) => ({ ...p, [t]: true }));
    setTimeout(() => setCopied((p) => ({ ...p, [t]: false })), 2200);
  };

  const dlTxt = (t, name) => {
    let txt = "";
    if (t === "r") txt = result.tailored_resume;
    else if (t === "c") txt = result.cover_letter;
    else
      txt = (result.application_answers || [])
        .map((a) => `Q: ${a.question}\n\nA: ${a.answer}`)
        .join("\n\n---\n\n");
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([txt], { type: "text/plain" }));
    a.download = name;
    a.click();
  };

  const dlDocx = async (t, fname) => {
    const txt =
      t === "r"
        ? result.tailored_resume
        : t === "c"
        ? result.cover_letter
        : (result.application_answers || [])
            .map((a) => `Q: ${a.question}\n\nA: ${a.answer}`)
            .join("\n\n---\n\n");
    const {
      Document,
      Packer,
      Paragraph,
      TextRun,
      AlignmentType,
      UnderlineType,
    } = window.docx;
    const lines = txt.split("\n");
    const KNOWN_HEADERS = [
      "experience",
      "education",
      "skills",
      "summary",
      "objective",
      "profile",
      "certifications",
      "awards",
      "publications",
      "projects",
      "volunteer",
      "languages",
      "interests",
      "references",
      "achievements",
      "accomplishments",
      "work experience",
      "professional experience",
      "technical skills",
      "core competencies",
      "professional summary",
      "career summary",
      "key skills",
      "additional information",
      "employment history",
      "qualifications",
    ];
    const norm = (s) =>
      s
        .toLowerCase()
        .replace(/[^a-z\s]/g, "")
        .trim();
    const isSectionHead = (line) => {
      const tr = line.trim();
      if (!tr || tr.length > 72) return false;
      const n = norm(tr);
      if (
        KNOWN_HEADERS.some(
          (h) => n === h || n === h + "s" || n.startsWith(h + " ")
        )
      )
        return true;
      if (
        tr === tr.toUpperCase() &&
        tr.length > 2 &&
        /[A-Z]/.test(tr) &&
        !/@|http|\d{3}/.test(tr)
      )
        return true;
      return false;
    };
    const isContactLine = (line) =>
      /[@|•·]|\d{3}[\s.\-]\d{3,4}|\blinkedin\b|\bgithub\b|\bwww\b|http/i.test(
        line
      );
    const isBulletLine = (line) => /^[\s]*[•\-\*◦▸►✓–]\s+\S/.test(line);
    const isRoleHeader = (line) => {
      const tr = line.trim();
      return (
        tr.length > 0 &&
        tr.length < 90 &&
        !tr.endsWith(".") &&
        (/\|\s|\s\|\s*/.test(tr) ||
          (/\d{4}/.test(tr) && tr.length < 80 && !isBulletLine(tr)))
      );
    };
    const children = [];
    let nameWritten = false,
      contactZoneDone = false;

    if (t === "r") {
      for (let i = 0; i < lines.length; i++) {
        const raw = lines[i];
        const tr = raw.trim();
        if (!tr) {
          children.push(
            new Paragraph({
              children: [new TextRun("")],
              spacing: { after: 60 },
            })
          );
          continue;
        }
        if (!nameWritten) {
          nameWritten = true;
          children.push(
            new Paragraph({
              children: [
                new TextRun({
                  text: tr,
                  bold: true,
                  size: 36,
                  font: "Calibri",
                  color: "111827",
                }),
              ],
              spacing: { after: 80 },
            })
          );
          continue;
        }
        if (!contactZoneDone) {
          if (isContactLine(tr) || i <= 4) {
            if (isContactLine(tr)) {
              children.push(
                new Paragraph({
                  children: [
                    new TextRun({
                      text: tr,
                      size: 18,
                      font: "Calibri",
                      color: "4B5563",
                    }),
                  ],
                  spacing: { after: 40 },
                })
              );
              continue;
            }
          }
          contactZoneDone = true;
        }
        contactZoneDone = true;
        if (isSectionHead(tr)) {
          children.push(
            new Paragraph({
              children: [
                new TextRun({
                  text: tr.toUpperCase(),
                  bold: true,
                  size: 22,
                  font: "Calibri",
                  color: "111827",
                }),
              ],
              spacing: { before: 300, after: 100 },
              border: {
                bottom: { color: "9CA3AF", size: 6, space: 4, style: "single" },
              },
            })
          );
          continue;
        }
        if (isRoleHeader(tr)) {
          const parts = tr.split(/\s*\|\s*/);
          const runs = parts.flatMap((p, idx) => [
            new TextRun({
              text: p.trim(),
              bold: idx === 0,
              italics: idx === parts.length - 1 && /\d{4}|present/i.test(p),
              size: 21,
              font: "Calibri",
              color: "111827",
            }),
            ...(idx < parts.length - 1
              ? [
                  new TextRun({
                    text: " · ",
                    size: 21,
                    font: "Calibri",
                    color: "9CA3AF",
                  }),
                ]
              : []),
          ]);
          children.push(
            new Paragraph({
              children: runs,
              spacing: { before: 140, after: 60 },
            })
          );
          continue;
        }
        if (isBulletLine(tr)) {
          children.push(
            new Paragraph({
              children: [
                new TextRun({
                  text: tr.replace(/^[•\-\*◦▸►✓–]\s+/, ""),
                  size: 20,
                  font: "Calibri",
                  color: "1F2937",
                }),
              ],
              bullet: { level: 0 },
              spacing: { after: 60 },
            })
          );
          continue;
        }
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: tr,
                size: 20,
                font: "Calibri",
                color: "1F2937",
              }),
            ],
            spacing: { after: 80 },
          })
        );
      }
    } else {
      let buf = [];
      const flush = () => {
        if (!buf.length) return;
        const fullText = buf.join(" ");
        const isSign =
          /^(sincerely|regards|best|warm regards|yours|thank you)[,.]?$/i.test(
            fullText.trim()
          );
        const isQ = fullText.startsWith("Q:");
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: fullText,
                size: 22,
                font: "Calibri",
                color: "1F2937",
                bold: isSign || isQ,
              }),
            ],
            spacing: { after: isSign ? 320 : 180 },
          })
        );
        buf = [];
      };
      for (const line of lines) {
        const tr = line.trim();
        if (!tr) {
          flush();
          continue;
        }
        buf.push(tr);
      }
      flush();
    }

    const doc = new Document({
      styles: {
        default: {
          document: { run: { font: "Calibri", size: 20, color: "1F2937" } },
        },
      },
      sections: [
        {
          properties: {
            page: {
              margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 },
            },
          },
          children,
        },
      ],
    });
    const blob = await Packer.toBlob(doc);
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = fname + ".docx";
    a.click();
  };

  const score = result?.ats_score || 0;
  const scoreColor = score >= 80 ? C.green : score >= 60 ? C.yellow : C.red;
  const CIRC = 251.3;
  const dashOffset = atsAnim ? CIRC - (score / 100) * CIRC : CIRC;
  const hasQAs = result?.application_answers?.length > 0;

  const togBtn = (active) => ({
    flex: 1,
    padding: "6px",
    border: "none",
    borderRadius: 5,
    fontSize: ".78rem",
    cursor: "pointer",
    background: active ? C.accent : "transparent",
    color: active ? "#fff" : C.muted,
    fontFamily: "inherit",
    fontWeight: active ? 500 : 400,
    transition: "all .18s",
  });
  const outBtn = (hi) => ({
    padding: "6px 14px",
    borderRadius: 6,
    border: `1px solid ${hi ? C.green : C.border}`,
    background: C.bg,
    color: hi ? C.green : "#c9d1d9",
    fontSize: ".76rem",
    cursor: "pointer",
    fontFamily: "inherit",
    display: "flex",
    alignItems: "center",
    gap: 5,
    transition: "all .18s",
  });

  const tabs = [
    ["r", "📄 Resume"],
    ["c", "✉️ Cover Letter"],
    ...(hasQAs ? [["q", "❓ Application Q&A"]] : []),
  ];

  return (
    <div
      style={{
        background: C.bg,
        minHeight: "100vh",
        color: C.text,
        fontFamily: '-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif',
      }}
    >
      <style>{`@keyframes spin{to{transform:rotate(360deg)}} textarea:focus,input:focus{outline:none;border-color:#6c63ff!important} textarea,input{transition:border-color .2s} ::-webkit-scrollbar{width:5px} ::-webkit-scrollbar-track{background:#0d1117} ::-webkit-scrollbar-thumb{background:#30363d;border-radius:3px}`}</style>

      {/* Header */}
      <div
        style={{
          background: "linear-gradient(135deg,#13162a,#0d1117)",
          borderBottom: `1px solid ${C.border}`,
          padding: "18px 32px",
          display: "flex",
          alignItems: "center",
          gap: 14,
        }}
      >
        <div
          style={{
            width: 38,
            height: 38,
            background: "linear-gradient(135deg,#6c63ff,#a855f7)",
            borderRadius: 10,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            fontSize: 17,
            flexShrink: 0,
          }}
        >
          ✦
        </div>
        <div>
          <div
            style={{
              fontSize: "1.35rem",
              fontWeight: 800,
              letterSpacing: "-.03em",
              background: "linear-gradient(135deg,#e6edf3,#a5b4fc)",
              WebkitBackgroundClip: "text",
              WebkitTextFillColor: "transparent",
            }}
          >
            Hyrd
          </div>
          <div style={{ fontSize: ".78rem", color: C.muted, marginTop: 1 }}>
            Tailor your resume, cover letter &amp; application — naturally,
            strategically, undetectably.
          </div>
        </div>
      </div>

      <div style={{ maxWidth: 1180, margin: "0 auto", padding: "28px 24px" }}>
        {/* Input grid */}
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "1fr 1fr",
            gap: 20,
            marginBottom: 20,
          }}
        >
          <div
            style={{
              background: C.card,
              border: `1px solid ${C.border}`,
              borderRadius: 12,
              padding: 18,
            }}
          >
            <div
              style={{
                fontSize: ".72rem",
                fontWeight: 600,
                textTransform: "uppercase",
                letterSpacing: ".08em",
                color: C.muted,
                marginBottom: 10,
              }}
            >
              📋 Job Description
            </div>
            <textarea
              value={jd}
              onChange={(e) => setJd(e.target.value)}
              placeholder="Paste the full job posting — title, responsibilities, requirements, everything..."
              style={{
                width: "100%",
                background: C.bg,
                border: `1px solid ${C.border}`,
                borderRadius: 8,
                color: C.text,
                padding: 12,
                fontSize: ".85rem",
                lineHeight: 1.65,
                resize: "vertical",
                minHeight: 240,
                fontFamily: "inherit",
              }}
            />
          </div>
          <div
            style={{
              background: C.card,
              border: `1px solid ${C.border}`,
              borderRadius: 12,
              padding: 18,
            }}
          >
            <div
              style={{
                fontSize: ".72rem",
                fontWeight: 600,
                textTransform: "uppercase",
                letterSpacing: ".08em",
                color: C.muted,
                marginBottom: 10,
              }}
            >
              📄 Your Resume
            </div>
            <div
              style={{
                display: "flex",
                background: C.bg,
                borderRadius: 7,
                padding: 3,
                gap: 3,
                marginBottom: 10,
              }}
            >
              <button
                style={togBtn(mode === "paste")}
                onClick={() => setMode("paste")}
              >
                ✏️ Paste Text
              </button>
              <button
                style={togBtn(mode === "file")}
                onClick={() => setMode("file")}
              >
                📎 Upload File
              </button>
            </div>
            {mode === "paste" ? (
              <textarea
                value={resText}
                onChange={(e) => setResText(e.target.value)}
                placeholder="Paste your current resume here..."
                style={{
                  width: "100%",
                  background: C.bg,
                  border: `1px solid ${C.border}`,
                  borderRadius: 8,
                  color: C.text,
                  padding: 12,
                  fontSize: ".85rem",
                  lineHeight: 1.65,
                  resize: "vertical",
                  minHeight: 200,
                  fontFamily: "inherit",
                }}
              />
            ) : (
              <label
                onDragOver={(e) => {
                  e.preventDefault();
                  setDrag(true);
                }}
                onDragLeave={() => setDrag(false)}
                onDrop={(e) => {
                  e.preventDefault();
                  setDrag(false);
                  handleFile(e.dataTransfer.files[0]);
                }}
                style={{
                  border: `2px dashed ${
                    fileName ? C.green : drag ? C.accent : C.border
                  }`,
                  borderRadius: 8,
                  minHeight: 200,
                  display: "flex",
                  flexDirection: "column",
                  alignItems: "center",
                  justifyContent: "center",
                  gap: 10,
                  cursor: "pointer",
                  background: fileName
                    ? "rgba(63,185,80,.05)"
                    : drag
                    ? "rgba(108,99,255,.06)"
                    : "transparent",
                  transition: "all .2s",
                }}
              >
                <input
                  type="file"
                  accept=".pdf,.docx"
                  style={{ display: "none" }}
                  onChange={(e) => handleFile(e.target.files[0])}
                />
                <div style={{ fontSize: "2rem" }}>{fileName ? "✅" : "📎"}</div>
                <div
                  style={{
                    fontSize: ".82rem",
                    color: C.muted,
                    textAlign: "center",
                    lineHeight: 1.5,
                  }}
                >
                  {fileName ? (
                    <span style={{ color: C.green, fontWeight: 600 }}>
                      {fileName} — parsed ✓
                    </span>
                  ) : (
                    <>
                      <strong style={{ color: C.text }}>
                        Drop your file here
                      </strong>{" "}
                      or click to browse
                      <br />
                      <span style={{ fontSize: ".72rem", color: "#484f58" }}>
                        PDF and DOCX supported
                      </span>
                    </>
                  )}
                </div>
              </label>
            )}
          </div>
        </div>

        {/* Application Questions */}
        <div
          style={{
            background: C.card,
            border: `1px solid ${C.border}`,
            borderRadius: 12,
            padding: 18,
            marginBottom: 20,
          }}
        >
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              marginBottom: 14,
            }}
          >
            <div>
              <div
                style={{
                  fontSize: ".72rem",
                  fontWeight: 600,
                  textTransform: "uppercase",
                  letterSpacing: ".08em",
                  color: C.muted,
                }}
              >
                ❓ Application Questions{" "}
                <span
                  style={{
                    color: "#484f58",
                    fontWeight: 400,
                    textTransform: "none",
                    letterSpacing: 0,
                  }}
                >
                  (optional)
                </span>
              </div>
              <div
                style={{ fontSize: ".75rem", color: "#484f58", marginTop: 3 }}
              >
                Add any questions from the job application — cover letter
                prompts, "Why us?", behavioural questions, etc.
              </div>
            </div>
            <button
              onClick={addQ}
              style={{
                padding: "6px 14px",
                borderRadius: 7,
                border: `1px dashed ${C.accent}`,
                background: "rgba(108,99,255,.08)",
                color: "#a5b4fc",
                fontSize: ".78rem",
                cursor: "pointer",
                fontFamily: "inherit",
                whiteSpace: "nowrap",
                flexShrink: 0,
              }}
            >
              + Add Question
            </button>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {questions.map((q, i) => (
              <div
                key={i}
                style={{ display: "flex", gap: 8, alignItems: "center" }}
              >
                <input
                  value={q}
                  onChange={(e) => updateQ(i, e.target.value)}
                  placeholder={`Question ${
                    i + 1
                  } — e.g. "Why do you want to work here?"`}
                  style={{
                    flex: 1,
                    background: C.bg,
                    border: `1px solid ${C.border}`,
                    borderRadius: 8,
                    color: C.text,
                    padding: "10px 12px",
                    fontSize: ".84rem",
                    fontFamily: "inherit",
                  }}
                />
                {questions.length > 1 && (
                  <button
                    onClick={() => removeQ(i)}
                    style={{
                      width: 30,
                      height: 30,
                      borderRadius: 6,
                      border: `1px solid ${C.border}`,
                      background: "transparent",
                      color: C.muted,
                      cursor: "pointer",
                      fontSize: "1rem",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      flexShrink: 0,
                    }}
                  >
                    ×
                  </button>
                )}
              </div>
            ))}
          </div>
        </div>

        {/* Generate */}
        <button
          disabled={loading}
          onClick={generate}
          style={{
            width: "100%",
            padding: 13,
            background: loading
              ? "#21262d"
              : "linear-gradient(135deg,#6c63ff,#a855f7)",
            border: "none",
            borderRadius: 10,
            color: "#fff",
            fontSize: ".95rem",
            fontWeight: 600,
            cursor: loading ? "not-allowed" : "pointer",
            marginBottom: 28,
            letterSpacing: ".02em",
            fontFamily: "inherit",
            opacity: loading ? 0.7 : 1,
          }}
        >
          {loading ? loadMsg : "✦  Tailor My Resume & Generate Cover Letter"}
        </button>

        {loading && (
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              gap: 12,
              padding: 48,
              color: C.muted,
              fontSize: ".88rem",
            }}
          >
            <div
              style={{
                width: 18,
                height: 18,
                border: `2px solid ${C.border}`,
                borderTopColor: C.accent,
                borderRadius: "50%",
                animation: "spin .75s linear infinite",
              }}
            />
            {loadMsg}
          </div>
        )}

        {error && (
          <div
            style={{
              color: C.red,
              fontSize: ".82rem",
              padding: 12,
              background: "rgba(248,81,73,.08)",
              border: "1px solid rgba(248,81,73,.2)",
              borderRadius: 8,
              marginBottom: 16,
            }}
          >
            {error}
          </div>
        )}

        {/* Output */}
        {result && (
          <>
            {/* ATS */}
            <div
              style={{
                background: "linear-gradient(135deg,#13162a,#161b22)",
                border: `1px solid ${C.border}`,
                borderRadius: 12,
                padding: "22px 24px",
                marginBottom: 20,
                display: "flex",
                alignItems: "flex-start",
                gap: 28,
              }}
            >
              <div
                style={{
                  position: "relative",
                  width: 96,
                  height: 96,
                  flexShrink: 0,
                }}
              >
                <svg
                  width="96"
                  height="96"
                  viewBox="0 0 96 96"
                  style={{ transform: "rotate(-90deg)" }}
                >
                  <circle
                    cx="48"
                    cy="48"
                    r="40"
                    fill="none"
                    stroke={C.border}
                    strokeWidth="9"
                  />
                  <circle
                    cx="48"
                    cy="48"
                    r="40"
                    fill="none"
                    stroke={scoreColor}
                    strokeWidth="9"
                    strokeLinecap="round"
                    strokeDasharray="251.3"
                    strokeDashoffset={dashOffset}
                    style={{
                      transition:
                        "stroke-dashoffset 1.2s cubic-bezier(.4,0,.2,1)",
                    }}
                  />
                </svg>
                <div
                  style={{
                    position: "absolute",
                    inset: 0,
                    display: "flex",
                    flexDirection: "column",
                    alignItems: "center",
                    justifyContent: "center",
                  }}
                >
                  <div
                    style={{
                      fontSize: "1.55rem",
                      fontWeight: 700,
                      color: scoreColor,
                    }}
                  >
                    {score}
                  </div>
                  <div
                    style={{
                      fontSize: ".6rem",
                      color: C.muted,
                      textTransform: "uppercase",
                      letterSpacing: ".06em",
                    }}
                  >
                    ATS Score
                  </div>
                </div>
              </div>
              <div style={{ flex: 1 }}>
                <div
                  style={{
                    fontSize: ".95rem",
                    fontWeight: 600,
                    marginBottom: 2,
                  }}
                >
                  {result.ats_confidence} Match Confidence
                </div>
                <div
                  style={{
                    fontSize: ".78rem",
                    color: C.muted,
                    marginBottom: 14,
                    lineHeight: 1.5,
                  }}
                >
                  {result.ats_summary}
                </div>
                <div
                  style={{ display: "flex", flexDirection: "column", gap: 7 }}
                >
                  {Object.entries(result.ats_breakdown || {}).map(
                    ([lbl, pct]) => {
                      const bc =
                        pct >= 80 ? C.green : pct >= 60 ? C.yellow : C.red;
                      return (
                        <div
                          key={lbl}
                          style={{
                            display: "flex",
                            alignItems: "center",
                            gap: 10,
                            fontSize: ".75rem",
                          }}
                        >
                          <div style={{ width: 130, color: "#8b949e" }}>
                            {lbl}
                          </div>
                          <div
                            style={{
                              flex: 1,
                              height: 5,
                              background: C.border,
                              borderRadius: 3,
                              overflow: "hidden",
                            }}
                          >
                            <div
                              style={{
                                height: "100%",
                                borderRadius: 3,
                                background: bc,
                                width: atsAnim ? pct + "%" : "0%",
                                transition:
                                  "width 1.1s cubic-bezier(.4,0,.2,1)",
                              }}
                            />
                          </div>
                          <div
                            style={{
                              width: 32,
                              textAlign: "right",
                              fontWeight: 500,
                            }}
                          >
                            {pct}%
                          </div>
                        </div>
                      );
                    }
                  )}
                </div>
                <div
                  style={{
                    display: "flex",
                    flexWrap: "wrap",
                    gap: 5,
                    marginTop: 13,
                  }}
                >
                  {(result.keywords_added || []).map((k) => (
                    <span
                      key={k}
                      style={{
                        padding: "3px 9px",
                        borderRadius: 20,
                        fontSize: ".7rem",
                        fontWeight: 500,
                        background: "rgba(108,99,255,.12)",
                        border: "1px solid rgba(108,99,255,.25)",
                        color: "#a5b4fc",
                      }}
                    >
                      + {k}
                    </span>
                  ))}
                </div>
              </div>
            </div>

            {/* Tabs */}
            <div style={{ display: "flex", gap: 3 }}>
              {tabs.map(([key, label]) => (
                <button
                  key={key}
                  onClick={() => setTab(key)}
                  style={{
                    padding: "9px 18px",
                    border: `1px solid ${C.border}`,
                    borderBottom: "none",
                    background: tab === key ? C.card : C.bg,
                    color: tab === key ? C.text : C.muted,
                    borderRadius: "8px 8px 0 0",
                    cursor: "pointer",
                    fontSize: ".82rem",
                    fontWeight: 500,
                    fontFamily: "inherit",
                  }}
                >
                  {label}
                </button>
              ))}
            </div>
            <div
              style={{
                background: C.card,
                border: `1px solid ${C.border}`,
                borderRadius: "0 8px 8px 8px",
                padding: 20,
              }}
            >
              {/* Resume & Cover Letter tabs */}
              {["r", "c"].map((t) => (
                <div key={t} style={{ display: tab === t ? "block" : "none" }}>
                  <div
                    style={{
                      display: "flex",
                      gap: 8,
                      justifyContent: "flex-end",
                      marginBottom: 13,
                    }}
                  >
                    <button style={outBtn(copied[t])} onClick={() => cp(t)}>
                      {copied[t] ? "✓ Copied!" : "📋 Copy"}
                    </button>
                    <button
                      style={outBtn(false)}
                      onClick={() =>
                        dlTxt(
                          t,
                          t === "r" ? "tailored_resume.txt" : "cover_letter.txt"
                        )
                      }
                    >
                      ⬇️ .txt
                    </button>
                    <button
                      style={outBtn(false)}
                      onClick={() =>
                        dlDocx(
                          t,
                          t === "r" ? "Tailored_Resume" : "Cover_Letter"
                        )
                      }
                    >
                      📝 .docx
                    </button>
                  </div>
                  <div
                    style={{
                      background: C.bg,
                      border: `1px solid ${C.border}`,
                      borderRadius: 8,
                      padding: 18,
                      fontSize: ".83rem",
                      lineHeight: 1.8,
                      whiteSpace: "pre-wrap",
                      color: C.text,
                      maxHeight: 520,
                      overflowY: "auto",
                    }}
                  >
                    {t === "r" ? result.tailored_resume : result.cover_letter}
                  </div>
                </div>
              ))}

              {/* Q&A tab */}
              {hasQAs && (
                <div style={{ display: tab === "q" ? "block" : "none" }}>
                  <div
                    style={{
                      display: "flex",
                      gap: 8,
                      justifyContent: "flex-end",
                      marginBottom: 13,
                    }}
                  >
                    <button style={outBtn(copied.q)} onClick={() => cp("q")}>
                      {copied.q ? "✓ Copied!" : "📋 Copy All"}
                    </button>
                    <button
                      style={outBtn(false)}
                      onClick={() => dlTxt("q", "application_answers.txt")}
                    >
                      ⬇️ .txt
                    </button>
                    <button
                      style={outBtn(false)}
                      onClick={() => dlDocx("q", "Application_Answers")}
                    >
                      📝 .docx
                    </button>
                  </div>
                  <div
                    style={{
                      display: "flex",
                      flexDirection: "column",
                      gap: 16,
                    }}
                  >
                    {(result.application_answers || []).map((qa, i) => (
                      <div
                        key={i}
                        style={{
                          background: C.bg,
                          border: `1px solid ${C.border}`,
                          borderRadius: 8,
                          padding: 18,
                        }}
                      >
                        <div
                          style={{
                            fontSize: ".75rem",
                            fontWeight: 600,
                            textTransform: "uppercase",
                            letterSpacing: ".06em",
                            color: C.accent,
                            marginBottom: 8,
                          }}
                        >
                          Question {i + 1}
                        </div>
                        <div
                          style={{
                            fontSize: ".88rem",
                            fontWeight: 600,
                            color: C.text,
                            marginBottom: 10,
                            lineHeight: 1.5,
                          }}
                        >
                          {qa.question}
                        </div>
                        <div
                          style={{
                            fontSize: ".84rem",
                            color: "#c9d1d9",
                            lineHeight: 1.8,
                            borderTop: `1px solid ${C.border}`,
                            paddingTop: 10,
                          }}
                        >
                          {qa.answer}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </>
        )}
      </div>
    </div>
  );
}
