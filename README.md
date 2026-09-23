# Excel Automation & POS Reconciliation Streamlit Dashboard

A high-performance, low-memory, production-grade multi-page Streamlit application designed for deploying batch Excel automation and POS reconciliation workflows on cloud infrastructure (e.g. Streamlit Community Cloud) under a strict **1.0 GB to 2.7 GB RAM limit**.

---

## 🏗️ Architecture & Low-Memory Strategy

Traditionally, executing openpyxl scripts across 50+ Excel workbooks eagerly loads XML cell trees into Python memory, rapidly consuming 1.5–3+ GB of RAM and causing Out-Of-Memory (OOM) crashes in cloud containers.

This dashboard eliminates memory bottlenecks via the following architectural pillars:

1. **Chunked Disk Spooling (`core/memory.py`)**:
   - Streamlit `UploadedFile` streams are spooled to disk in 64 KB chunks using `tempfile.NamedTemporaryFile`.
   - Prevents memory spikes from holding dozens of binary file buffers in RAM simultaneously.
2. **Iterative Batch Processing (`core/gear/engine.py`)**:
   - Workbooks are processed **one-by-one**.
   - Each source workbook is opened, extracted, immediately closed, and its temporary file unlinked before the next file is spooled.
   - Deterministic garbage collection (`gc.collect()`) is triggered after every file.
3. **Selective & Read-Only Parsing**:
   - Source workbooks are inspected with `data_only=True` and closed immediately in `finally` blocks.
   - Only the single destination master workbook is held open for modifications.
4. **Zero Aggressive Caching**:
   - No `@st.cache_data` on binary file streams or openpyxl DOMs.
   - Session state only retains lightweight scalar metadata and temporary output paths.
5. **Form Encapsulation (`st.form`)**:
   - Every tool interface is encapsulated in `st.form` to prevent unnecessary script reruns when dragging and dropping 50+ files.

---

## 🛠️ Integrated Automation Suite

The application is structured into four dedicated tools accessible via Streamlit's native multi-page sidebar:

| Page | Tool Name | Description & Capabilities |
|---|---|---|
| `Home.py` | **System Dashboard** | Real-time process RAM monitor, health metrics, and tool directory. |
| `pages/1_⛽_Gear_Automation.py` | **Station Fuel Gear Aggregator** | Batch processes 50+ daily station workbooks into the monthly Gear master. Performs fuzzy station matching, updates dipping & meter sales, injects running average formulas, and computes master PMS/AGO total columns. |
| `pages/2_📋_Cashbook_Extractor.py` | **Cashbook Loop Extractor** | Parses reconciliation cashbooks, locates report blocks, and extracts active terminal numbers for `SGR`, `OBX`, `AGS`, and `TA` into compact JSON for the Moniepoint Chrome Extension. |
| `pages/3_📥_EOD_Cashbook_Sync.py` | **EOD Cashbook Sync** | Filters Moniepoint EOD transaction logs by date, aggregates amounts by normalized POS, writes into the `CASHBOOK` column, recalculates difference formulas (`CASHBOOK - BANK`), and appends a styled audit worksheet. |
| `pages/4_🏦_Bank_Statement_Loader.py` | **Bank Statement POS Loader** | Consumes Chrome Extension JSON exports, matches terminal rows, populates the `BANK STATEMENT` column with currency formatting, and attaches detailed cell comments for failed records. |

---

## 📁 Repository Layout

```text
STREAMLIT_APP/
├── .streamlit/
│   └── config.toml               # Max upload limit (500MB), server headless, theme
├── Home.py                       # Main executive dashboard & diagnostics
├── pages/
│   ├── 1_⛽_Gear_Automation.py    # Fuel station 50+ batch aggregator
│   ├── 2_📋_Cashbook_Extractor.py # Chrome extension loop JSON builder
│   ├── 3_📥_EOD_Cashbook_Sync.py  # Moniepoint EOD transaction synchronizer
│   └── 4_🏦_Bank_Statement_Loader.py # POS Bot JSON to Bank Statement loader
├── core/                         # Pure stateless Python domain logic
│   ├── memory.py                 # Disk spooling, cleanup, and RAM monitor
│   ├── gear/
│   │   ├── config.py             # GearConfig dataclass & row validation
│   │   ├── utils.py              # Tank normalization, number parser
│   │   ├── extractors.py         # Tank tables & meter sales extractors
│   │   ├── matcher.py            # Fuzzy station block matcher & column resolver
│   │   ├── formulas.py           # Divisor incrementing & total formula generator
│   │   └── engine.py             # Iterative batch aggregation pipeline
│   └── pos_recon/
│       ├── common.py             # POS regex normalizer, date & money parsers
│       ├── cashbook_extractor.py # POS loop grouping logic
│       ├── eod_extractor.py      # EOD source reader & audit sheet builder
│       └── bank_loader.py        # Chrome extension JSON parser & comment injector
├── utils/                        # Streamlit UI & presentation helpers
│   ├── ui_components.py          # Metric cards, memory badge, download buttons
│   └── audit_viewer.py           # DataFrame log table formatters
├── tests/                        # Automated unit and integration test suite
│   ├── test_memory.py            # Spooling and cleanup verification
│   ├── test_gear.py              # Formulas and Gear batch integration
│   ├── test_pos_recon.py         # POS normalizer, cashbook, EOD, and bank tests
│   └── run_all_tests.py          # Standalone test runner
├── requirements.txt              # Production dependencies
└── README.md                     # Documentation
```

---

## 🚀 Getting Started

### 1. Installation

Create and activate a virtual environment, then install dependencies:

```bash
# Create venv
python3 -m venv .venv
source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### 2. Run Automated Verification Tests

Verify that all memory management, fuzzy matching, and reconciliation engines pass:

```bash
python tests/run_all_tests.py
```

### 3. Launch the Application

```bash
streamlit run Home.py
```

Open your browser at `http://localhost:8501`.

---

## ☁️ Cloud Deployment (Streamlit Community Cloud)

1. **Push to GitHub**:
   Push this directory to your GitHub repository.
2. **Deploy via Streamlit Cloud**:
   - Go to [share.streamlit.io](https://share.streamlit.io).
   - Select your repository and branch.
   - Set **Main file path** to `Home.py`.
3. **Upload Limit Configuration**:
   The included `.streamlit/config.toml` already configures:
   ```toml
   [server]
   maxUploadSize = 500
   ```
   This ensures users can drag and drop 50+ station Excel files in a single batch without triggering file size rejections.
