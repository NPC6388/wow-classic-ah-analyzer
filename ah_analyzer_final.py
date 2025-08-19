import re
import csv
import os
from datetime import datetime
from collections import defaultdict

def convert_price_to_gold(price_str):
    """Convert price from copper to gold.silver.copper format"""
    try:
        copper_total = int(price_str)
        
        # Convert copper to gold, silver, copper
        # 1 gold = 100 silver = 10000 copper
        gold = copper_total // 10000
        remaining_copper = copper_total % 10000
        silver = remaining_copper // 100
        copper = remaining_copper % 100
        
        # Return as decimal gold (gold + silver/100 + copper/10000)
        return gold + (silver / 100) + (copper / 10000)
    except:
        return 0

def format_price_wow(price_gold):
    """Format price for display as WoW format: XgYsZc"""
    try:
        # Convert decimal gold back to copper for formatting
        copper_total = int(price_gold * 10000)
        gold = copper_total // 10000
        remaining_copper = copper_total % 10000
        silver = remaining_copper // 100
        copper = remaining_copper % 100
        return f"{gold}g {silver}s {copper}c"
    except:
        return "0g 0s 0c"

def format_price_copper_to_wow(copper_total):
    """Format copper directly to WoW format: XgYsZc"""
    try:
        copper_total = int(copper_total)
        gold = copper_total // 10000
        remaining_copper = copper_total % 10000
        silver = remaining_copper // 100
        copper = remaining_copper % 100
        return f"{gold}g {silver}s {copper}c"
    except:
        return "0g 0s 0c"

def parse_auc_stat_stddev(file_path):
    """Parse Auc-Stat-StdDev.lua file to get market price data"""
    print(f"Processing StdDev file: {file_path}")
    
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except Exception as e:
        print(f"Error reading StdDev file: {e}")
        return {}
    
    market_prices = {}
    
    # Look for item data in format: [12640] = "0:price1;price2;price3;..."
    pattern = r'\[(\d+)\]\s*=\s*"[^:]*:([^"]+)"'
    
    matches = re.finditer(pattern, content)
    
    for match in matches:
        item_id = int(match.group(1))
        price_data = match.group(2)
        
        # Parse the price values (semicolon separated)
        price_parts = price_data.split(';')
        if len(price_parts) >= 1:
            try:
                # Use the most recent price (last entry) as market price
                recent_price_copper = int(price_parts[-1])
                market_price_gold = convert_price_to_gold(recent_price_copper)
                market_prices[item_id] = market_price_gold
                    
            except (ValueError, IndexError):
                continue
    
    print(f"Found market price data for {len(market_prices)} items")
    return market_prices

def parse_auc_stat_histogram(file_path):
    """Parse Auc-Stat-Histogram.lua file to get times seen and market price data"""
    print(f"Processing histogram file: {file_path}")
    
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except Exception as e:
        print(f"Error reading histogram file: {e}")
        return {}, {}
    
    times_seen = {}
    market_prices = {}
    
    # Look for item data in format: ["12640"] = "0@percentile!percentile!price!count!bins;histogram_data"
    pattern = r'\["(\d+)"\]\s*=\s*"[^!]+!([^!]+)!(\d+)!(\d+)!'
    
    matches = re.finditer(pattern, content)
    
    for match in matches:
        item_id = int(match.group(1))
        market_price_copper = int(match.group(3))
        count = int(match.group(4))
        
        # Convert market price from copper to gold
        market_price_gold = convert_price_to_gold(market_price_copper)
        
        # Store the highest count found for each item (in case there are multiple entries)
        if item_id not in times_seen or times_seen[item_id] < count:
            times_seen[item_id] = count
            market_prices[item_id] = market_price_gold
    
    print(f"Found histogram data for {len(times_seen)} items")
    return times_seen, market_prices

def parse_auc_stat_simple(file_path):
    """Parse Auc-Stat-Simple.lua file to get times seen data"""
    print(f"Processing stat file: {file_path}")
    
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except Exception as e:
        print(f"Error reading stat file: {e}")
        return {}
    
    times_seen = {}
    
    # Look for item data in format: ["12640"] = "0@something;count;prices..."
    pattern = r'\["(\d+)"\]\s*=\s*"[^;]+;(\d+);'
    
    matches = re.finditer(pattern, content)
    
    for match in matches:
        item_id = int(match.group(1))
        count = int(match.group(2))
        
        # Store the highest count found for each item (in case there are multiple entries)
        if item_id not in times_seen or times_seen[item_id] < count:
            times_seen[item_id] = count
    
    print(f"Found times seen data for {len(times_seen)} items")
    return times_seen

def parse_auctioneer_data(file_path):
    """Parse Auctioneer scan data from Lua file"""
    print(f"Processing: {file_path}")
    
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        print(f"File size: {len(content)} characters")
    except Exception as e:
        print(f"Error reading file: {e}")
        return []
    
    items = []
    
    # Look for the ropes section with the return format
    ropes_match = re.search(r'\["ropes"\]\s*=\s*\{([^}]+(?:\{[^}]*\}[^}]*)*)\}', content, re.DOTALL)
    
    if ropes_match:
        ropes_content = ropes_match.group(1)
        print(f"Found ropes section, length: {len(ropes_content)}")
        
        # Look for the return statement with the actual data
        return_match = re.search(r'return\s*\{([^}]+(?:\{[^}]*\}[^}]*)*)\}', ropes_content, re.DOTALL)
        
        if return_match:
            data_content = return_match.group(1)
            print(f"Found return data, length: {len(data_content)}")
            
            # Parse individual item entries
            # Pattern: {"item_link",level,quality,count,nil,buyout,bid,time,seller_name,...
            item_pattern = r'\{([^}]+(?:\{[^}]*\}[^}]*)*)\}'
            
            matches = re.finditer(item_pattern, data_content, re.DOTALL)
            
            for match in matches:
                try:
                    item_data = match.group(1)
                    
                    # Extract item link and name
                    item_link_match = re.search(r'\|Hitem:(\d+):[^|]*\|h\[([^\]]+)\]\|h\|r', item_data)
                    if not item_link_match:
                        continue
                    
                    item_id = int(item_link_match.group(1))
                    item_name = item_link_match.group(2)
                    
                    # Split the rest of the data by commas
                    parts = item_data.split(',')
                    
                    # Find the numeric values after the item link
                    # Skip the item link part and look for numbers
                    numeric_parts = []
                    for part in parts:
                        part = part.strip().strip('"')
                        if part.isdigit():
                            numeric_parts.append(int(part))
                    
                    if len(numeric_parts) >= 10:
                        level = numeric_parts[0]
                        quality = numeric_parts[1]
                        count = numeric_parts[2]
                        buyout_price = numeric_parts[4]  # Skip the nil value
                        bid_price = numeric_parts[5]
                        time_left = numeric_parts[6]
                        # Position 14 in the data structure contains scan frequency
                        scan_frequency = numeric_parts[9] if len(numeric_parts) > 9 else 1
                        
                        # Extract seller name (it's usually after the time value)
                        seller_match = re.search(r'(\d+),(\d+),(\d+),(\d+),(\d+),(\d+),(\d+),"([^"]+)"', item_data)
                        if seller_match:
                            seller_name = seller_match.group(7)
                        else:
                            seller_name = "Unknown"
                        
                        # Convert price to gold using our pricing logic
                        buyout_gold = convert_price_to_gold(buyout_price)
                        bid_gold = convert_price_to_gold(bid_price)
                        
                        items.append({
                            'item_id': item_id,
                            'item_name': item_name,
                            'level': level,
                            'quality': quality,
                            'count': count,
                            'buyout_price_copper': buyout_price,
                            'buyout_price_gold': buyout_gold,
                            'bid_price_copper': bid_price,
                            'bid_price_gold': bid_gold,
                            'time_left': time_left,
                            'seller_name': seller_name,
                            'scan_frequency': scan_frequency
                        })
                        
                except Exception as e:
                    print(f"Error parsing item: {e}")
                    continue
            
            print(f"Found {len(items)} items in ropes section")
    
    # If no items found in ropes, try alternative parsing
    if len(items) == 0:
        print("Trying alternative parsing...")
        
        # Look for any item-like patterns in the entire file
        alt_pattern = r'\|Hitem:(\d+):[^|]*\|h\[([^\]]+)\]\|h\|r'
        
        matches = re.finditer(alt_pattern, content, re.DOTALL)
        
        for match in matches:
            try:
                item_id = int(match.group(1))
                item_name = match.group(2)
                
                # Look for the data after this item link
                item_start = match.start()
                item_end = match.end()
                
                # Get the next 200 characters after the item link
                next_data = content[item_end:item_end+200]
                
                # Look for numeric values
                numbers = re.findall(r'(\d+)', next_data)
                
                if len(numbers) >= 10:
                    level = int(numbers[0])
                    quality = int(numbers[1])
                    count = int(numbers[2])
                    buyout_price = int(numbers[4])  # Skip nil
                    bid_price = int(numbers[5])
                    time_left = int(numbers[6])
                    # Position 14 in data structure contains scan frequency
                    scan_frequency = int(numbers[9]) if len(numbers) > 9 else 1
                    
                    # Convert price to gold using our pricing logic
                    buyout_gold = convert_price_to_gold(buyout_price)
                    bid_gold = convert_price_to_gold(bid_price)
                    
                    items.append({
                        'item_id': item_id,
                        'item_name': item_name,
                        'level': level,
                        'quality': quality,
                        'count': count,
                        'buyout_price_copper': buyout_price,
                        'buyout_price_gold': buyout_gold,
                        'bid_price_copper': bid_price,
                        'bid_price_gold': bid_gold,
                        'time_left': time_left,
                        'seller_name': "Unknown",
                        'scan_frequency': scan_frequency
                    })
                    
            except Exception as e:
                continue
        
        print(f"Alternative parsing found {len(items)} items")
    
    print(f"Total items found: {len(items)}")
    return items

def analyze_arbitrage(horde_items, alliance_items, horde_times_seen=None, alliance_times_seen=None, horde_market_prices=None, alliance_market_prices=None):
    """Analyze cross-faction arbitrage opportunities"""
    print("Analyzing arbitrage opportunities...")
    
    # Group items by name
    horde_by_name = defaultdict(list)
    alliance_by_name = defaultdict(list)
    
    for item in horde_items:
        horde_by_name[item['item_name']].append(item)
    
    for item in alliance_items:
        alliance_by_name[item['item_name']].append(item)
    
    # Find items that exist on both factions
    common_items = set(horde_by_name.keys()) & set(alliance_by_name.keys())
    print(f"Found {len(common_items)} items on both factions")
    
    arbitrage_opportunities = []
    
    for item_name in common_items:
        horde_items_list = horde_by_name[item_name]
        alliance_items_list = alliance_by_name[item_name]
        
        # Calculate average prices (for backup if market price not available)
        horde_avg_price = sum(item['buyout_price_gold'] for item in horde_items_list) / len(horde_items_list)
        alliance_avg_price = sum(item['buyout_price_gold'] for item in alliance_items_list) / len(alliance_items_list)
        
        # Get market prices from histogram data (prefer this over average)
        item_id = horde_items_list[0]['item_id'] if horde_items_list else alliance_items_list[0]['item_id']
        horde_market_price = horde_market_prices.get(item_id, horde_avg_price) if horde_market_prices else horde_avg_price
        alliance_market_price = alliance_market_prices.get(item_id, alliance_avg_price) if alliance_market_prices else alliance_avg_price
        
        # Find minimum buyout prices (cheapest listing on each faction)
        horde_min_price = min(item['buyout_price_gold'] for item in horde_items_list)
        alliance_min_price = min(item['buyout_price_gold'] for item in alliance_items_list)
        
        # Find price difference between minimum prices
        price_diff = abs(horde_min_price - alliance_min_price)
        
        # Determine which faction is cheaper based on minimum prices
        if horde_min_price < alliance_min_price:
            cheaper_faction = "Horde"
        else:
            cheaper_faction = "Alliance"
        
        # Get times seen from histogram files
        horde_total_scans = horde_times_seen.get(item_id, 0) if horde_times_seen else 0
        alliance_total_scans = alliance_times_seen.get(item_id, 0) if alliance_times_seen else 0
        
        arbitrage_opportunities.append({
            'item_name': item_name,
            'horde_market_price': horde_market_price,
            'alliance_market_price': alliance_market_price,
            'horde_buyout_price': horde_min_price,
            'alliance_buyout_price': alliance_min_price,
            'price_difference': price_diff,
            'cheaper_faction': cheaper_faction,
            'horde_scan_count': horde_total_scans,
            'alliance_scan_count': alliance_total_scans
        })
    
    # Sort by price difference (largest differences first)
    arbitrage_opportunities.sort(key=lambda x: x['price_difference'], reverse=True)
    
    return arbitrage_opportunities

def generate_csv_report(horde_items, alliance_items, arbitrage_opportunities):
    """Generate CSV reports"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Generate comprehensive report
    csv_filename = f"ah_analysis_{timestamp}.csv"
    
    with open(csv_filename, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        
        # Write header
        writer.writerow([
            'Item Name', 'Times Seen (Horde, Alliance)', 'Horde Buyout Price', 
            'Alliance Buyout Price', 'Price Difference', 'Horde Market Price', 
            'Alliance Market Price', 'Cheaper Faction'
        ])
        
        # Write arbitrage opportunities
        for opp in arbitrage_opportunities:
            # Round price difference to gold only (no silver/copper)
            price_diff_gold = int(round(opp['price_difference']))
            
            writer.writerow([
                opp['item_name'],
                f"{opp['horde_scan_count']}, {opp['alliance_scan_count']}",
                format_price_wow(opp['horde_buyout_price']),
                format_price_wow(opp['alliance_buyout_price']),
                f"{price_diff_gold}g",
                format_price_wow(opp['horde_market_price']),
                format_price_wow(opp['alliance_market_price']),
                opp['cheaper_faction']
            ])
    
    print(f"Generated CSV report: {csv_filename}")
    
    # Generate separate faction reports
    horde_csv = f"horde_bargains_{timestamp}.csv"
    with open(horde_csv, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['Item Name', 'Buyout Price (Gold)', 'Count', 'Seller'])
        
        # Group Horde items by name and find lowest prices
        horde_by_name = defaultdict(list)
        for item in horde_items:
            horde_by_name[item['item_name']].append(item)
        
        horde_bargains = []
        for name, items in horde_by_name.items():
            min_price_item = min(items, key=lambda x: x['buyout_price_gold'])
            horde_bargains.append({
                'name': name,
                'price': min_price_item['buyout_price_gold'],
                'count': min_price_item['count'],
                'seller': min_price_item['seller_name']
            })
        
        horde_bargains.sort(key=lambda x: x['price'])
        for item in horde_bargains[:100]:
            writer.writerow([item['name'], format_price_wow(item['price']), item['count'], item['seller']])
    
    alliance_csv = f"alliance_bargains_{timestamp}.csv"
    with open(alliance_csv, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['Item Name', 'Buyout Price (Gold)', 'Count', 'Seller'])
        
        # Group Alliance items by name and find lowest prices
        alliance_by_name = defaultdict(list)
        for item in alliance_items:
            alliance_by_name[item['item_name']].append(item)
        
        alliance_bargains = []
        for name, items in alliance_by_name.items():
            min_price_item = min(items, key=lambda x: x['buyout_price_gold'])
            alliance_bargains.append({
                'name': name,
                'price': min_price_item['buyout_price_gold'],
                'count': min_price_item['count'],
                'seller': min_price_item['seller_name']
            })
        
        alliance_bargains.sort(key=lambda x: x['price'])
        for item in alliance_bargains[:100]:
            writer.writerow([item['name'], format_price_wow(item['price']), item['count'], item['seller']])
    
    print(f"Generated Horde bargains: {horde_csv}")
    print(f"Generated Alliance bargains: {alliance_csv}")
    
    return csv_filename, horde_csv, alliance_csv

def main():
    """Main function"""
    print("WoW Classic AH Analyzer - Final Version")
    print("=" * 50)
    
    # File paths
    horde_path = r"C:\Program Files (x86)\World of Warcraft\_classic_era_\WTF\Account\LUCIA1\SavedVariables\auc-scandata.lua"
    alliance_path = r"C:\Program Files (x86)\World of Warcraft\_classic_era_\WTF\Account\51718250#1\SavedVariables\auc-scandata.lua"
    
    # Histogram file paths (contains times seen)
    horde_histogram_path = r"C:\Program Files (x86)\World of Warcraft\_classic_era_\WTF\Account\LUCIA1\SavedVariables\Auc-Stat-Histogram.lua"
    alliance_histogram_path = r"C:\Program Files (x86)\World of Warcraft\_classic_era_\WTF\Account\51718250#1\SavedVariables\Auc-Stat-Histogram.lua"
    
    # StdDev file paths (contains market prices)
    horde_stddev_path = r"C:\Program Files (x86)\World of Warcraft\_classic_era_\WTF\Account\LUCIA1\SavedVariables\Auc-Stat-StdDev.lua"
    alliance_stddev_path = r"C:\Program Files (x86)\World of Warcraft\_classic_era_\WTF\Account\51718250#1\SavedVariables\Auc-Stat-StdDev.lua"
    
    print(f"Horde data path: {horde_path}")
    print(f"Alliance data path: {alliance_path}")
    print(f"Horde histogram path: {horde_histogram_path}")
    print(f"Alliance histogram path: {alliance_histogram_path}")
    print(f"Horde StdDev path: {horde_stddev_path}")
    print(f"Alliance StdDev path: {alliance_stddev_path}")
    print()
    
    # Check if files exist
    if not os.path.exists(horde_path):
        print(f"ERROR: Horde scan data not found at {horde_path}")
        return
    
    if not os.path.exists(alliance_path):
        print(f"ERROR: Alliance scan data not found at {alliance_path}")
        return
    
    # Extract times seen data from histogram files
    print("Extracting Horde histogram data...")
    horde_times_seen, _ = parse_auc_stat_histogram(horde_histogram_path)
    
    print("Extracting Alliance histogram data...")
    alliance_times_seen, _ = parse_auc_stat_histogram(alliance_histogram_path)
    
    # Extract market price data from StdDev files
    print("Extracting Horde market price data...")
    horde_market_prices = parse_auc_stat_stddev(horde_stddev_path)
    
    print("Extracting Alliance market price data...")
    alliance_market_prices = parse_auc_stat_stddev(alliance_stddev_path)
    
    # Extract auction data
    print("Extracting Horde auction data...")
    horde_items = parse_auctioneer_data(horde_path)
    
    print("Extracting Alliance auction data...")
    alliance_items = parse_auctioneer_data(alliance_path)
    
    if not horde_items and not alliance_items:
        print("No auction data found in either file!")
        print("This might indicate the data format is different than expected.")
        print("Please check that the files contain Auctioneer scan data.")
        return
    
    print(f"\nTotal items found:")
    print(f"Horde: {len(horde_items)}")
    print(f"Alliance: {len(alliance_items)}")
    
    if len(horde_items) == 0 and len(alliance_items) == 0:
        print("No items found. The data format may be different than expected.")
        return
    
    # Analyze arbitrage opportunities
    arbitrage_opportunities = analyze_arbitrage(horde_items, alliance_items, horde_times_seen, alliance_times_seen, horde_market_prices, alliance_market_prices)
    
    if len(arbitrage_opportunities) == 0:
        print("No arbitrage opportunities found.")
        print("This could mean:")
        print("- No items exist on both factions")
        print("- Price differences are too small")
        print("- Data parsing needs adjustment")
    
    # Generate reports
    main_csv, horde_csv, alliance_csv = generate_csv_report(horde_items, alliance_items, arbitrage_opportunities)
    
    print(f"\nAnalysis complete!")
    print(f"Main report: {main_csv}")
    print(f"Top 100 Horde bargains: {horde_csv}")
    print(f"Top 100 Alliance bargains: {alliance_csv}")
    
    # Open the main CSV file
    try:
        os.startfile(main_csv)
        print("Opened main report in Excel")
    except:
        print(f"Please open {main_csv} manually in Excel")

if __name__ == "__main__":
    main()

