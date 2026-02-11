# Audit: Ironclad Assessment Rendering & Chunking
## Architecture for Universal Consistency

To move from a functional prototype to an "ironclad" production system, we must shift the philosophy from **Text Manipulation** (Regex/Strings) to **Semantic Processing** (DOM/Structure). This audit outlines the path to ensuring 100% consistent rendering regardless of the source file's complexity.

---

### 1. The Core Problem: The "Flattening" Trap
Currently, the system converts files to HTML and then uses Regex to "guess" where questions begin. This fails on:
*   **Multi-column layouts:** OCR often reads left-to-right across columns, mixing text.
*   **Tables:** Reading a table cell-by-cell without context is confusing.
*   **Nested Lists:** (e.g., 1.a.i.) Regex often breaks on the sub-bullets rather than the question.

---

### 2. Strategy: The "Golden Intermediate" Format
Instead of rendering the raw (and messy) HTML from the OCR/Conversion process, we should adopt a **Structured Intermediate Representation (SIR)**.

**Proposed Pipeline:**
1.  **Source Conversion:** PDF/Doc → Raw HTML (Current).
2.  **Semantic Normalization:** Instead of just sanitizing, we map elements to a strict schema:
    *   `QuestionBlock` (Container)
    *   `QuestionText` (Paragraphs)
    *   `OptionList` (Answer Choices)
    *   `Asset` (Image/Table)
3.  **JSON-First Chunking:** Chunking happens on the JSON structure, not the string.

---

### 3. Improving OCR & Layout Awareness
The current Google Drive OCR is a "blind" text extractor.
*   **Gemini Vision Integration:** For PDFs with complex layouts (columns, sidebars), we should use the Gemini Pro Vision model to "describe the layout" or "convert this image to clean Markdown."
*   **Advantage:** Gemini understands that a sidebar is separate from the main flow. It can return structured Markdown which is significantly easier to chunk than messy HTML.

---

### 4. Semantic DOM Chunking (Replacing Regex)
Instead of `CHUNK_SPLIT_REGEX`, we should implement a **Walking Parser**:
*   **Logic:** Iterate through the DOM tree. 
*   **Boundary Detection:** A "Chunk" is defined as a node that contains a specific pattern (e.g., a paragraph starting with a number) **plus** all its sibling nodes until the next boundary is found.
*   **Benefit:** This preserves images and tables *within* the chunk they belong to. Currently, if an image sits between two paragraphs, regex might split it away from its context.

---

### 5. Handling Tables and Complex Objects
Tables are the "final boss" of Read-Aloud.
*   **Linearization:** We need a `tableToSpeech` utility that converts a table into a descriptive string: *"Table titled 'Climate Data': Row 1, Column 1 is Year, Column 2 is Temp..."*
*   **Rendering:** On the student side, the table remains a visual table, but the "Audio Chunk" contains this linearized description.

---

### 6. Universal CSS "Reset" for Assessments
To ensure perfect rendering, the Student Portal should inject a "Shadow DOM" or a strict CSS reset specifically for the `html-assessment-container`.
*   **Fluid Typography:** Use `rem` and `ch` units to ensure text wraps identically on a Chromebook, iPad, or Phone.
*   **Standardized Spacing:** Force all converted `<p>` tags to have identical margins, overriding any artifacts from the original Word/PDF file.

---

### 7. Implementation Roadmap: The "Ironclad" Upgrades

#### Phase 1: Markdown Intermediate (Low Effort)
Convert the raw HTML to **Markdown** before processing. Markdown is structured, removes 90% of HTML "noise," and has clear boundaries for lists and headers.

#### Phase 2: Boundary Object Mapping (Medium Effort)
Change the data structure in the spreadsheet from:
`[{text: "...", audioUrl: "..."}]`
to
`[{id: "q1", type: "question", content: "...", children: ["opt_a", "opt_b"], audioUrl: "..."}]`
This allows the UI to render "Smart Blocks" instead of just "Injected HTML."

#### Phase 3: Multimodal Layout Recovery (High Effort)
For PDFs, use Gemini 1.5 Pro to "Understand and OCR" the document. Send the PDF as an image/file to the model with a prompt: *"Convert this assessment into a structured JSON format with Question, Options, and Context fields."* This bypasses the limitations of traditional OCR entirely.

---

### Summary of Recommendations
1.  **Abandon String-based Regex splitting** in favor of a DOM-traversal parser.
2.  **Adopt Markdown** as the intermediate format to strip Google Docs/PDF styling artifacts.
3.  **Implement Layout Detection** via Gemini Vision for complex "non-linear" PDF assessments.
4.  **Linearize Tables** for the TTS engine while maintaining the visual grid for the student.
