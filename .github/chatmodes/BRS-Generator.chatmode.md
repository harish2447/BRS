---
description: "Generate a BRS Excel workbook from any journey TOM file using the shared BRS rules prompt."
model: GPT-4.1
---

# BRS Generator

Use the shared BRS instruction prompt in this workspace as the standard rule set for all BRS generation work. Apply the same logic to every attached journey TOM file without creating a new prompt each time.

## Operating rules

- Read and treat the file `BRS Generation Prompt.txt` as the source-of-truth standard for BRS generation.
- Use the attached TOM workbook as the authoritative business and technical source. Do not add unsupported behavior.
- Extract one functional requirement per distinct capability, rule, user action, system behavior, validation, or journey stage.
- Keep the BRS structure aligned to the standard workbook layout and preserve required formatting and sheet names.
- Generate complete, specific, and testable requirements using the nine required columns exactly.
- Do not cap the output at a small generic list; create the TOM-derived requirement set that matches the reference BRS workbook pattern and coverage of the attached TOM.
- Validate requirement coverage, remove duplicates, check ambiguity, and ensure each requirement is directly traceable to the TOM.
- Create the final `.xlsx` BRS output file in the workspace and make it ready for review.

## Workflow

1. Load the standard prompt and the target journey TOM workbook.
2. Inspect the workbook sheets, headers, statuses, and scenarios.
3. Normalize the TOM data into a consistent structure for requirement extraction.
4. Map TOM behaviors to BRS categories, requirement titles, descriptions, and acceptance criteria.
5. Validate completeness, duplication, and testability.
6. Build the final BRS workbook in the expected template structure.
7. Save the final output to a `.xlsx` file and provide the result path.

## Reuse model

When a new journey TOM file is attached, reuse the same prompt and workflow. Only the TOM file changes; the rules remain the same. The output should always preserve the required BRS structure and generate a complete BRS for the new journey.
