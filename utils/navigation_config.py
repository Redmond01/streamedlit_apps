from __future__ import annotations

# ==============================================================================
# 🛠️ CENTRAL SIDEBAR NAVIGATION CONFIGURATION
# ==============================================================================
# You can freely and safely change the "title", "icon", or section names below.
# Changing these values updates the sidebar navigation labels immediately
# without altering any filenames, underlying scripts, or core business logic!
# ==============================================================================

NAVIGATION_CONFIG = {
    "Overview": [
        {
            "target": "home",
            "title": "System Overview",
            "icon": "⚡",
            "default": True,
        }
    ],
    "Fuel Operations": [
        {
            "target": "pages/1_⛽_Gear_Automation.py",
            "title": "Gear Automation",             # <-- Edit your display title here!
            "icon": "⛽",
        },
        {
            "target": "pages/5_📊_Daily_Sales_Aggregator.py",
            "title": "Daily Sales Aggregator",      # <-- Edit your display title here!
            "icon": "📊",
        },
    ],
    "POS & Bank Reconciliation": [
        {
            "target": "pages/2_📋_Cashbook_Extractor.py",
            "title": "EOD JSON Extractor",         # <-- Edit your display title here!
            "icon": "📋",
        },
        {
            "target": "pages/3_📥_EOD_Cashbook_Sync.py",
            "title": "EOD API Loader",              # <-- Edit your display title here!
            "icon": "📥",
        },
        {
            "target": "pages/4_🏦_Bank_Statement_Loader.py",
            "title": "Bank Statement POS Loader",   # <-- Edit your display title here!
            "icon": "🏦",
        },
    ],
}
