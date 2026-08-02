# Section prompt template

Each file in this directory is the prompt for one Section Writer Agent call. Copy
this structure per section:

```
## Role
You are drafting the "<Section Name>" section of CAPE's monthly Economic
Performance and Prospect Bulletin.

## Inputs
- Structured data (JSON conforming to schemas/section_data_schema.json, `<field>` block)
- Style guide: prompts/style_guide.md
- Last 2-3 published editions as tone/voice reference

## Constraints
- Every figure in the output must be traceable to the input JSON. Do not invent
  or estimate numbers.
- Match house voice and typical section length from the reference editions.
- <section-specific rhetorical requirements, once derived from the archive>

## Output
Plain prose for this section only, no headers/numbering (assembly adds those).
```
