# SKILL.md

## Name
Productivity Document, Spreadsheet, CSV, and Automation Assistant

## When to use this skill
Use this skill when the task involves one or more of these work products:
- Drafting, rewriting, summarising, or checking documents.
- Creating structured reports from notes, transcripts, PDFs, forms, or research.
- Cleaning, analysing, validating, or summarising spreadsheets and CSV files.
- Turning repeated office tasks into a checklist, script, template, or automation plan.

Do not use this skill for legal, medical, financial, procurement, HR discipline, or other high-risk decisions without a named human reviewer.

## Inputs to collect first
- **Goal:** What should the finished output help the user decide or do?
- **Audience:** Who will read or use the output?
- **Source files:** Document, spreadsheet, CSV, report, transcript, email thread, or notes.
- **Required format:** Memo, report, table, slide outline, CSV, spreadsheet formulas, checklist, or automation steps.
- **Constraints:** Deadline, word count, required columns, tone, style guide, privacy limits, and citation requirements.
- **Verification source:** The file, row, page, link, or authority used to confirm each important claim.

## Standard workflow
1. Identify the output type: document, spreadsheet, report, CSV, or automation.
2. Inspect the source before writing. Note missing fields, unclear instructions, duplicate records, and possible private data.
3. Create a short plan with the exact sections, columns, formulas, or automation steps needed.
4. Produce the first draft or transformed dataset.
5. Verify facts, numbers, formulas, row counts, dates, names, and citations against the source.
6. Flag assumptions and unresolved questions instead of hiding uncertainty.
7. Return the final output in the requested format with a brief quality checklist.

## Document and report rules
- Keep the main message visible in the first paragraph or executive summary.
- Preserve names, dates, figures, and policy references exactly unless the user asks for anonymisation.
- Cite source pages, sections, rows, or links for factual claims.
- Separate facts, analysis, recommendations, and assumptions.
- Use tables when comparing options, action items, owners, deadlines, risks, or metrics.

## Spreadsheet and CSV rules
- Never silently drop rows or columns. Report before/after row counts.
- Keep a copy of the original field names unless renaming is requested.
- Mark missing, duplicate, invalid, or outlier values.
- Use explicit formulas and explain what each calculated column means.
- For CSV output, keep one header row, consistent delimiters, and no merged cells.

## Automation rules
- Start with the manual process: trigger, inputs, decision points, output, owner, and exception path.
- Prefer simple automations first: template, checklist, spreadsheet formula, email rule, or scheduled script.
- Include failure handling, logging, and a human approval step for sensitive outputs.
- Never automate sending, deleting, approving, or publishing without explicit human confirmation.

## Output quality checklist
- [ ] Output matches the requested format.
- [ ] All required sections or columns are present.
- [ ] Important facts and numbers were checked against the source.
- [ ] Assumptions and gaps are listed.
- [ ] Private or sensitive data was not exposed unnecessarily.
- [ ] The user can act on the result without extra cleanup.

## Example prompt
```text
Role: You are a productivity automation analyst.
Task: Convert this CSV of training attendance into a verified summary report and action-item table.
Context: The audience is a regional program manager. The source CSV has participant name, office, email, attendance_status, score, and comments.
Format:
1. 150-word executive summary.
2. Table of attendance counts by office.
3. List of participants needing follow-up.
4. Data quality issues found.
Constraints:
- Do not invent missing scores.
- Do not expose email addresses in the summary.
- Verify all totals against the CSV row count.
```

## Escalation
Stop and ask a human reviewer when:
- The source data contains personal, confidential, or regulated information.
- The output will be used for a formal decision about a person, budget, procurement, or compliance matter.
- The data has contradictions that change the recommendation.
- The user asks to bypass review, remove audit trails, or hide uncertainty.
