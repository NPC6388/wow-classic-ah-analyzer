# WoW Classic Cross-Faction Auction House Analyzer

A local Python tool for analyzing World of Warcraft Classic auction house data across factions using Auctioneer addon scan data.

## Features

- **Cross-Faction Price Comparison**: Compare prices between Horde and Alliance auction houses
- **Automated Data Reading**: Reads directly from Auctioneer SavedVariables files
- **Price Delta Analysis**: Calculate price differences and identify arbitrage opportunities
- **CSV Export**: Generate detailed reports with item comparisons
- **Proper WoW Currency Parsing**: Correctly handles copper/silver/gold price format

## Requirements

- Python 3.6+
- Auctioneer addon installed on both Horde and Alliance characters
- Recent auction house scans on both factions
- All required files must exist: `auc-scandata.lua`, `Auc-Stat-Histogram.lua`, and `Auc-Stat-StdDev.lua`

## Installation

1. Ensure Python is installed on your system
2. Clone or download this repository
3. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Run auction house scans with Auctioneer on both your Horde and Alliance characters
2. Run the analyzer:
```bash
python ah_analyzer_final.py
```

The script will:
- Read Horde auction data from: `C:\Program Files (x86)\World of Warcraft\_classic_era_\WTF\Account\LUCIA1\SavedVariables\auc-scandata.lua`
- Read Alliance auction data from: `C:\Program Files (x86)\World of Warcraft\_classic_era_\WTF\Account\51718250#1\SavedVariables\auc-scandata.lua`
- Read Horde times seen data from: `C:\Program Files (x86)\World of Warcraft\_classic_era_\WTF\Account\LUCIA1\SavedVariables\Auc-Stat-Histogram.lua`
- Read Alliance times seen data from: `C:\Program Files (x86)\World of Warcraft\_classic_era_\WTF\Account\51718250#1\SavedVariables\Auc-Stat-Histogram.lua`
- Read Horde market price data from: `C:\Program Files (x86)\World of Warcraft\_classic_era_\WTF\Account\LUCIA1\SavedVariables\Auc-Stat-StdDev.lua`
- Read Alliance market price data from: `C:\Program Files (x86)\World of Warcraft\_classic_era_\WTF\Account\51718250#1\SavedVariables\Auc-Stat-StdDev.lua`
- Generate a CSV file with columns: Item Name, Times Seen (Horde, Alliance), Horde Buyout Price, Alliance Buyout Price, Price Difference, Horde Market Price, Alliance Market Price, Cheaper Faction
- Open the generated CSV file automatically

## Data Sources

The analyzer reads from three types of Auctioneer files:

### Auction Data (auc-scandata.lua)
- **Price Information**: Contains current auction listings with buyout prices
- **All prices stored as copper**: The raw data contains prices in total copper value
- **Conversion**: 1 gold = 100 silver = 10,000 copper
- **Example**: `1999800` copper = 199 gold, 98 silver, 0 copper = 199.98g

### Times Seen Data (Auc-Stat-Histogram.lua)
- **Times Seen**: Contains historical scan frequency data matching in-game Auctioneer display
- **Format**: Data stored as `["itemID"] = "0@percentile!percentile!price!count!bins;histogram..."`
- **Accuracy**: Times seen values match exactly what is shown in-game by Auctioneer addon

### Market Price Data (Auc-Stat-StdDev.lua)
- **Market Price**: Contains Auctioneer's calculated market prices based on historical data
- **Format**: Data stored as `[itemID] = "0:price1;price2;price3;...;mostrecent"`
- **Accuracy**: Uses the same pricing data that Auctioneer displays in-game for market values

## Output

The generated CSV file includes:
- **Item Name**: Display name of the item
- **Times Seen (Horde, Alliance)**: Number of times Auctioneer has scanned this item on each faction (matches in-game display, format: "144, 55")
- **Horde Buyout Price**: Cheapest listing price on Horde auction house (format: "1798g 19s 88c")
- **Alliance Buyout Price**: Cheapest listing price on Alliance auction house (format: "1261g 0s 0c")
- **Price Difference**: Absolute difference between cheapest prices, rounded to gold only (format: "537g")
- **Horde Market Price**: Auctioneer's calculated market price for Horde (format: "1799g 99s 87c")
- **Alliance Market Price**: Auctioneer's calculated market price for Alliance (format: "1199g 99s 99c")
- **Cheaper Faction**: Which faction has the lower minimum buyout price

## File Structure

```
wow-classic-ah-analyzer/
├── ah_analyzer_final.py     # Main analyzer script
├── requirements.txt         # Python dependencies
├── README.md               # This file
└── *.csv                   # Generated analysis reports
```

## Troubleshooting

- **File not found errors**: Ensure Auctioneer addon is installed and you've run recent scans
- **No data**: Make sure you've performed auction house scans on both factions recently
- **Path issues**: Verify the WoW installation path matches the default Classic Era location
- **Times seen showing 0**: Ensure the `Auc-Stat-Histogram.lua` files exist - these are generated after Auctioneer has collected statistical data over time
- **Market prices showing 0**: Ensure the `Auc-Stat-StdDev.lua` files exist - these contain historical pricing data used for market calculations
- **Incorrect values**: The script reads from the same statistical files that Auctioneer uses in-game, so times seen and market prices should match exactly

## Data Privacy

- All processing happens locally on your machine
- No data is sent to external servers
- Reads only from your local Auctioneer addon files