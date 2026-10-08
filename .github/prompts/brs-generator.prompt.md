# BRS Generator Prompt

Generate a Business Requirements Specification (BRS) Excel workbook from the supplied TOM file.

Use the shared prompt file `BRS Generation Prompt.txt` as the standard rule set and apply it consistently to every journey TOM file. Treat the TOM as the source of truth. Extract requirements only from the TOM and do not invent unsupported functionality.

Apply the following workflow to every TOM:

1. Read the common BRS instruction prompt and confirm the required output structure.
2. Inspect all relevant sheets, rows, columns, scenarios, validations, user roles, journey stages, outputs, and integrations.
3. Normalize the TOM data and identify distinct business capabilities and rules.
4. Derive one functional requirement for each distinct capability or rule.
5. Use the exact nine columns required by the BRS standard: Ref #, Requirement Category, Requirement Title, Requirement Description, Acceptance Criteria, Requirement SME, MoSCoW, System, Reference to Other Related Documents If Applicable.
6. Ensure each requirement is concise, specific, testable, and traceable to the TOM.
7. Remove duplicates and merge repeated scenarios where appropriate.
8. Validate coverage, ambiguity, and consistency before finalizing.
9. Build the final BRS workbook with the required template structure and save it as a real `.xlsx` workbook.

This prompt is reusable for any TOM file such as Journey 5&6, Journey 8, Journey 9, or similar future journey models. It must not stop at a short generic list; it should generate the full TOM-derived BRS set aligned to the reference workbook structure and the actual scenarios in the attached TOM.
