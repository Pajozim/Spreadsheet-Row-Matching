# Excel Semantic & SQL Search Engine

A lightweight desktop application built to accelerate row-matching workflows across spreadsheets by using structured SQL filtering and local vector-based semantic search.

---

## 🚀 Key Features

* **Asymmetric Semantic Search:** Leverages `multilingual-e5-small` with explicit instruction prefixes (`query:` for specification text, `passage:` for catalogue items) to accurately bridge long contract descriptions with concise database entries.
* **Hybrid Search Strategy:** Pairs deterministic SQL rules for hard parameters with vector similarity scoring for unstructured text.
* **Adaptive Compute Execution:** Automatically detects available local hardware to select either CUDA GPU acceleration or lightweight CPU processing without requiring manual re-configuration.
* **Local & Secure:** Fully contained local execution designed for WSL2, ensuring offline data privacy and zero cloud API dependency overhead.

---

## 🛣️ Roadmap & Planned Extensions

* [ ] **Domain Asset Extraction:** Custom parsing pipelines for supplementary data assets (`.txt`, `.json`, `.csv`) to ingest custom jargon and construction/industry abbreviation dictionaries.
* [ ] **Structured Unit & Metric Validation:** Post-processing validation layer to enforce strict metric compliance (e.g., matching length, area, or volume units before confirming row pairing).
* [ ] **Standardized Error Management:** Granular exception handling and structured reporting for missing database references or low-confidence matches.
* [ ] **Automated File Preservation:** Native duplication and safety-copying of original source files prior to write operations.

---

## ⚙️ Quick Start

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/Pajozim/Spreadsheet-Row-Matching.git](https://github.com/Pajozim/Spreadsheet-Row-Matching.git)
   cd Spreadsheet-Row-Matching
