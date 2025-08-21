# WoW Classic Cross-Faction Auction House Analyzer

A comprehensive Python tool for analyzing World of Warcraft Classic auction house data across factions using Auctioneer addon scan data. Identify profitable cross-faction arbitrage opportunities with accurate pricing and historical market data.

## ✨ Features

- **🔄 Cross-Faction Price Comparison**: Compare real-time buyout prices between Horde and Alliance auction houses
- **📊 Historical Market Analysis**: Incorporates Auctioneer's historical pricing and statistical data
- **💰 Buyout-Only Filtering**: Focuses on immediately purchasable items (excludes bid-only auctions)
- **📈 Arbitrage Opportunity Detection**: Identifies profitable cross-faction trading opportunities
- **📋 Excel Report Generation**: Professional Excel reports with interactive features
- **🧮 Per-Unit Price Calculator**: Built-in tool for calculating stack vs individual item pricing
- **📱 Dual Analysis View**: Shows both current buyout prices and historical market trends
- **🎯 Accurate Data Sourcing**: Uses the same statistical files as Auctioneer addon for consistency

## 📋 Requirements

- Python 3.6+
- **openpyxl** library for Excel output
- Auctioneer addon installed on both Horde and Alliance characters
- Recent auction house scans on both factions
- Required Auctioneer files: `auc-scandata.lua`, `Auc-Stat-Histogram.lua`, and `Auc-Stat-Simple.lua`

## 🚀 Installation

1. Ensure Python is installed on your system
2. Clone or download this repository
3. Install required dependencies:
```bash
pip install openpyxl
```

## 📖 Usage

1. **Run auction scans** with Auctioneer on both your Horde and Alliance characters
2. **Execute the analyzer**:
```bash
python ah_analyzer_final.py
```

The script will automatically:
- ✅ Read auction data from your WoW SavedVariables
- ✅ Process historical market data and statistics  
- ✅ Generate a comprehensive Excel report
- ✅ Open the report automatically for review

## 📁 Data Sources

### Current Auction Data (`auc-scandata.lua`)
- **Real-time prices**: Current auction listings with actual buyout prices
- **Buyout filtering**: Automatically excludes bid-only auctions
- **Accurate pricing**: Uses the correct buyout price positions (not bid prices)
- **Currency format**: All prices stored as copper, converted to gold display format

### Historical Statistics (`Auc-Stat-Simple.lua`)
- **Times Seen**: Number of times each item has been scanned by Auctioneer
- **Market Prices**: Auctioneer's calculated historical market values
- **Data Accuracy**: Matches exactly what Auctioneer displays in-game
- **Per-item pricing**: Market prices represent individual item costs, not stack prices

### Price Conversion
- **Currency**: 1 gold = 100 silver = 10,000 copper
- **Example**: `2,499,999 copper` = `249g 99s 99c`

## 📊 Excel Report Features

### Main Analysis Sheet
Contains detailed arbitrage opportunities with columns:

| Column | Description | Example |
|--------|-------------|---------|
| **Item Name** | Display name of the item | `Black Lotus` |
| **Times Seen (Horde, Alliance)** | Historical scan counts | `394, 1050` |
| **Horde Buyout Price** | Cheapest current Horde buyout | `47g 15s 0c` |
| **Alliance Buyout Price** | Cheapest current Alliance buyout | `35g 0s 0c` |
| **Price Difference** | Market price difference | `8g` |
| **Horde Market Price** | Historical market average | `26g 72s 46c` |
| **Alliance Market Price** | Historical market average | `30g 14s 30c` |
| **Cheaper Buyout** | Which faction has cheaper current price | `Alliance` |
| **Cheaper Historic** | Which faction has better historical price | `Horde` |

### Faction-Specific Bargain Sheets
- **Horde Bargains**: Top 100 underpriced items on Horde auction house
- **Alliance Bargains**: Top 100 underpriced items on Alliance auction house

### Interactive Per-Unit Calculator
Located below the main table:
- **Input Fields**: Stack size and total buyout price
- **Automatic Calculation**: Shows per-unit pricing in both copper and gold
- **Live Updates**: Recalculates instantly when inputs change
- **Professional Formatting**: Color-coded cells and clear instructions

**Calculator Example:**
- Stack Size: `20`
- Total Buyout: `50,000 copper`
- **Result**: `2,500 copper per unit (0.2500 gold per unit)`

## 🔧 Key Improvements

### ✅ Accurate Price Extraction
- **Fixed buyout vs bid confusion**: Now correctly shows buyout prices (higher) vs bid prices (lower)
- **Buyout-only focus**: Filters out bid-only auctions for immediate purchase opportunities
- **Proper position mapping**: Correctly extracts prices from Auctioneer data structure

### ✅ Enhanced Data Analysis
- **Historical context**: Uses Simple stat files for accurate times seen and market prices
- **Per-item pricing**: Market prices represent individual items, not stack totals
- **Smart filtering**: Only includes items with buyout options on both factions

### ✅ Professional Output
- **Excel format**: Replaced CSV with feature-rich Excel workbooks
- **Multiple sheets**: Organized data across different analysis views
- **Interactive tools**: Built-in calculator for price conversions
- **Visual formatting**: Professional styling with tables and color coding

## 📂 File Structure

```
wow-classic-ah-analyzer/
├── ah_analyzer_final.py           # Main analyzer script
├── README.md                      # This documentation
├── index.html                     # Web-based analyzer (legacy)
└── ah_analysis_YYYYMMDD_HHMMSS.xlsx # Generated Excel reports
```

## 🐛 Troubleshooting

### Common Issues
- **File not found**: Ensure Auctioneer addon is installed and you've run recent scans
- **No arbitrage opportunities**: Normal - indicates similar pricing across factions
- **Missing market data**: Run more auction scans to build Auctioneer's statistical database
- **Path issues**: Verify WoW installation matches default Classic Era location

### Data Validation
- **Times seen = 0**: `Auc-Stat-Simple.lua` files need time to accumulate scan data
- **Market prices = 0**: Historical data builds up over multiple scan sessions  
- **Buyout vs bid confusion**: Fixed - script now correctly identifies and shows buyout prices
- **Stack vs individual pricing**: Fixed - all prices represent per-item costs

## 🔒 Privacy & Security

- **100% Local Processing**: All analysis happens on your machine
- **No Network Requests**: No data sent to external servers
- **Read-Only Access**: Only reads from your Auctioneer addon files
- **No Game Modification**: Does not alter any WoW or addon files

## 🎯 Use Cases

### Cross-Faction Trading
- **Server Transfer Prep**: Identify which items to buy before transferring
- **Neutral Auction House**: Find profitable items for neutral AH trading
- **Market Analysis**: Understand faction-specific pricing trends

### Investment Opportunities  
- **Undervalued Items**: Find items selling below historical market value
- **Price Arbitrage**: Exploit temporary price differences between factions
- **Market Timing**: Compare current vs historical prices for buying decisions

## 📈 Example Analysis

**Black Lotus Arbitrage Opportunity:**
- Horde Current: `47.15g` vs Historic: `26.72g`
- Alliance Current: `35.00g` vs Historic: `30.14g`
- **Opportunity**: Buy on Alliance (cheaper current), potential profit: `12.15g per item`
- **Market Context**: Alliance historically cheaper, confirming trend

## 🤝 Contributing

This project was developed collaboratively with AI assistance. The codebase focuses on accuracy, reliability, and user-friendly analysis of WoW Classic auction house data.

---

**Last Updated**: August 21, 2025  
**Version**: 2.0 - Complete rewrite with Excel output, buyout filtering, and enhanced accuracy