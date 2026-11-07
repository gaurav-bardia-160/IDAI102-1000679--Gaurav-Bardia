"""
ShopImpact - Personal Carbon Accounting & Conscious Shopping Ledger
CRS: Artificial Intelligence | Course: Python Programming (Summative Assessment)
100% Pure Python Implementation (Zero HTML, CSS, or JavaScript)
Client-Side URL Compression: Survives Refresh & Zero Cross-Device Sync
"""

import base64
from datetime import date
import io
import json
import math
import random
import subprocess
import sys
import zlib
import altair as alt
import pandas as pd
from PIL import Image, ImageDraw
import streamlit as st

# Ensure Excel and PDF export libraries are available
try:
    import xlsxwriter  # noqa: F401
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "xlsxwriter", "-q"])
    import xlsxwriter  # noqa: F401

try:
    from fpdf import FPDF
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

# ==============================================================================
# 1. PAGE CONFIGURATION
# ==============================================================================
st.set_page_config(
    page_title="ShopImpact | Conscious Shopping Companion",
    page_icon="🌱",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==============================================================================
# 2. 108 PRODUCT CATEGORIES DATABASE (Calibrated for INR ₹)
# ==============================================================================
PRODUCT_CATALOG = {
    # --- 1. APPAREL & TEXTILES ---
    "Fast Fashion Polyester Dress / Top": ("Apparel & Textiles", 0.0048, "Switch to No Nasties, Doodlage, or thrift stores", 0.0005),
    "Conventional Cotton Jeans (New)": ("Apparel & Textiles", 0.0042, "Switch to Levi's SecondHand, Nudie Jeans, or recycled denim", 0.0006),
    "Virgin Wool Sweater / Coat": ("Apparel & Textiles", 0.0055, "Switch to Patagonia Recycled Wool or reclaimed wool knitwear", 0.0009),
    "Genuine Bovine Leather Jacket": ("Apparel & Textiles", 0.0068, "Switch to Malai (coconut bio-leather) or vintage leather", 0.0008),
    "Fast Fashion Synthetic Activewear": ("Apparel & Textiles", 0.0044, "Switch to Girlfriend Collective (recycled PET)", 0.0011),
    "Virgin Down Winter Parka": ("Apparel & Textiles", 0.0050, "Switch to 100% Recycled Down outerwear or PrimaLoft Bio", 0.0012),
    "Conventional Silk Blouse / Saree": ("Apparel & Textiles", 0.0039, "Switch to Ahimsa Peace Silk or Tencel Lyocell", 0.0010),
    "Nylon Raincoat / Windbreaker": ("Apparel & Textiles", 0.0041, "Switch to Econyl regenerated nylon outerwear", 0.0011),
    "Acrylic Knitwear / Beanie": ("Apparel & Textiles", 0.0046, "Switch to Boheco hemp knitwear or organic cotton", 0.0009),
    "Conventional Viscose / Kurta": ("Apparel & Textiles", 0.0038, "Switch to Lenzing EcoVero closed-loop viscose", 0.0010),
    "Organic Linen Apparel (New)": ("Apparel & Textiles", 0.0014, "Low-impact material; requires minimal irrigation", 0.0008),
    "Organic Hemp Clothing (New)": ("Apparel & Textiles", 0.0012, "Low-impact material; rain-fed hemp sequesters carbon", 0.0007),
    "Recycled Cotton Basics / T-Shirt": ("Apparel & Textiles", 0.0015, "Recycled cotton reduces dyeing water consumption by 90%", 0.0006),
    "Second-Hand / Thrift Store Apparel": ("Apparel & Textiles", 0.0004, "Circular benchmark; avoids 90%+ manufacturing emissions", 0.0004),
    "Peer-to-Peer Rent / Occasion Wear": ("Apparel & Textiles", 0.0007, "Renting ethnic wear avoids single-use purchases", 0.0005),
    "Upcycled / Deadstock Apparel": ("Apparel & Textiles", 0.0009, "Upcycled factory deadstock eliminates mill waste", 0.0006),

    # --- 2. FOOTWEAR & ACCESSORIES ---
    "Conventional Leather Formal Shoes": ("Footwear & Accessories", 0.0062, "Switch to chrome-free vegetable-tanned leather or resole", 0.0010),
    "Synthetic Running Sneakers (EVA)": ("Footwear & Accessories", 0.0047, "Switch to Neeman's or Allbirds sugarcane EVA soles", 0.0013),
    "Faux-Leather (PVC) Handbag": ("Footwear & Accessories", 0.0051, "Switch to Zouk, Phool Fleather, or recycled canvas bags", 0.0011),
    "Vulcanized Rubber Flip-Flops": ("Footwear & Accessories", 0.0036, "Switch to Paaduks (upcycled tire soles) or cork sandals", 0.0009),
    "Virgin Plastic Sunglasses": ("Footwear & Accessories", 0.0034, "Switch to bio-acetate frames or reclaimed marine plastic", 0.0008),
    "Newly Mined Gold / Diamond Jewelry": ("Footwear & Accessories", 0.0072, "Switch to 100% recycled hallmarked gold or lab-grown", 0.0009),
    "Newly Mined Silver Jewelry": ("Footwear & Accessories", 0.0045, "Switch to verified post-consumer recycled 925 silver", 0.0008),
    "Virgin Polyester Backpack / Luggage": ("Footwear & Accessories", 0.0040, "Switch to 100% rPET recycled bottle backpacks", 0.0010),
    "Recycled Ocean Plastic Footwear": ("Footwear & Accessories", 0.0013, "Neeman's ReLive shoes divert plastic from landfills", 0.0008),
    "Cork / Natural Crepe Footwear": ("Footwear & Accessories", 0.0011, "Harvested cork regenerates without felling trees", 0.0007),
    "Second-Hand / Pre-Owned Watch or Bag": ("Footwear & Accessories", 0.0005, "Pre-owned accessories retain craftsmanship at zero new carbon", 0.0004),
    "Shoe Cobbler Resole / Bag Repair": ("Footwear & Accessories", 0.0003, "Repairing soles uses <5% of the carbon of manufacturing", 0.0003),

    # --- 3. CONSUMER ELECTRONICS & TECH ---
    "Brand New Flagship Smartphone": ("Electronics & Tech", 0.0046, "Switch to Cashify / Apple Certified Refurbished", 0.0007),
    "Brand New Laptop / Ultrabook": ("Electronics & Tech", 0.0049, "Switch to certified refurbished business laptops", 0.0008),
    "Brand New Tablet / iPad": ("Electronics & Tech", 0.0044, "Switch to Apple Certified Refurbished tablets", 0.0007),
    "Desktop Gaming PC & Dedicated GPU": ("Electronics & Tech", 0.0053, "Purchase pre-owned GPUs and 80 PLUS Titanium PSUs", 0.0012),
    "4K / OLED Television (55-inch+)": ("Electronics & Tech", 0.0051, "Enable eco-dimming (cuts power 35%) or buy Open-Box", 0.0014),
    "Smartwatch / Fitness Wearable": ("Electronics & Tech", 0.0038, "Choose solar-assisted wearables or refurbished", 0.0008),
    "Wireless Bluetooth Earbuds (New)": ("Electronics & Tech", 0.0037, "Choose repairable battery earbuds (Fairbuds)", 0.0009),
    "Home Inkjet / Laser Printer (New)": ("Electronics & Tech", 0.0042, "Switch to Epson / Canon ink-tank bottle systems", 0.0010),
    "Video Game Console (PS5 / Xbox)": ("Electronics & Tech", 0.0045, "Buy pre-owned consoles and enable low-power rest mode", 0.0009),
    "Digital Mirrorless / DSLR Camera": ("Electronics & Tech", 0.0040, "Buy inspected pre-owned camera bodies (MPB / KEH)", 0.0006),
    "Single-Use Alkaline Batteries": ("Electronics & Tech", 0.0056, "Switch to Panasonic Eneloop NiMH rechargeable cells", 0.0007),
    "Virgin PVC Charging Cables": ("Electronics & Tech", 0.0035, "Switch to recycled rPET braided cables", 0.0009),
    "Certified Refurbished Smartphone": ("Electronics & Tech", 0.0007, "Avoids ~60 kg CO2e of raw earth mining", 0.0006),
    "Certified Refurbished Laptop": ("Electronics & Tech", 0.0008, "Avoids ~250 kg CO2e of manufacturing emissions", 0.0007),
    "Solar Power Bank / Eco Charger": ("Electronics & Tech", 0.0014, "Recycled-housing power banks with solar trickle charging", 0.0008),

    # --- 4. FOOD, GROCERIES & BEVERAGES ---
    "Red Meat / Mutton (Conventional)": ("Food & Groceries", 0.0078, "Swap for soya chunks, paneer, or masoor dal", 0.0009),
    "Farmed Warm-Water Prawns / Shrimp": ("Food & Groceries", 0.0064, "Swap for sustainable coastal bivalves or plant protein", 0.0008),
    "Processed Meat / Cured Sausages": ("Food & Groceries", 0.0042, "Swap for smoked tofu or tempeh (Hello Tempayy)", 0.0010),
    "Poultry / Broiler Chicken": ("Food & Groceries", 0.0031, "Swap for kabuli chana (chickpeas), rajma, or tofu", 0.0008),
    "Conventional Packet Dairy Milk": ("Food & Groceries", 0.0038, "Switch to oat/soy milk, or local loose milk in steel cans", 0.0010),
    "Processed Cheese Blocks / Slices": ("Food & Groceries", 0.0058, "Switch to fresh homemade chenna or plant-based spreads", 0.0012),
    "Dairy Butter & Heavy Cream": ("Food & Groceries", 0.0052, "Use cold-pressed mustard, sesame, or coconut oil", 0.0011),
    "Air-Freighted Imported Fruits": ("Food & Groceries", 0.0060, "Choose local seasonal Indian fruits (guava, banana, Chikoo)", 0.0007),
    "Sun-Grown Monoculture Coffee": ("Food & Groceries", 0.0045, "Switch to shade-grown Western Ghats agroforestry coffee", 0.0014),
    "Conventional Milk Chocolate": ("Food & Groceries", 0.0048, "Switch to 70%+ single-origin dark chocolate", 0.0015),
    "Bottled Mineral Water (PET)": ("Food & Groceries", 0.0040, "Use home RO/UV filtration + stainless steel bottle", 0.0003),
    "Imported Bottled Wine (Heavy Glass)": ("Food & Groceries", 0.0036, "Choose domestic Nashik vineyards or lightweight cans", 0.0012),
    "Canned Beverages / Aerated Drinks": ("Food & Groceries", 0.0025, "Switch to homemade nimbu pani or buttermilk in glass", 0.0005),
    "Packaged Chips & Palm-Oil Snacks": ("Food & Groceries", 0.0034, "Switch to roasted makhana (foxnuts) or bulk nuts", 0.0008),
    "Polished White Rice (Flooded Paddy)": ("Food & Groceries", 0.0033, "Swap for rain-fed Indian millets (Ragi, Jowar, Bajra)", 0.0007),
    "Plant-Based Milk (Oat / Soy / Millet)": ("Food & Groceries", 0.0010, "Oat milks require <15% of the water of dairy", 0.0008),
    "Dry Dal, Lentils, Chana & Rajma": ("Food & Groceries", 0.0006, "Pulses fix nitrogen naturally in soil", 0.0005),
    "Local Mandi Seasonal Vegetables": ("Food & Groceries", 0.0007, "Field-grown seasonal sabzi in cloth bags carries minimal carbon", 0.0005),
    "Unpackaged Kirana / Bulk Grains": ("Food & Groceries", 0.0006, "Loose grains in reusable cloth thailis removes plastic waste", 0.0005),
    "Rain-Fed Indigenous Millets": ("Food & Groceries", 0.0005, "Dryland crops requiring 70% less water than paddy rice", 0.0004),

    # --- 5. HOME, KITCHEN & CLEANING ---
    "Single-Use Kitchen Paper Towels": ("Home & Kitchen", 0.0038, "Switch to washable cotton kitchen dusters or bamboo tissue rolls", 0.0008),
    "Liquid Detergent (Plastic Jug)": ("Home & Kitchen", 0.0035, "Switch to laundry sheets or soapnut (reetha) powder", 0.0007),
    "Chemical Dishwash Liquid (Plastic)": ("Home & Kitchen", 0.0033, "Switch to coconut coir scrub pads + solid dishwash bars", 0.0007),
    "Virgin Plastic Garbage Bags": ("Home & Kitchen", 0.0041, "Segregate wet waste into a home compost pot (Khamba)", 0.0009),
    "Cling Film & Ziploc Freezer Pouches": ("Home & Kitchen", 0.0039, "Use stainless steel tiffin boxes or beeswax wraps", 0.0005),
    "Teflon (PTFE) Non-Stick Cookware": ("Home & Kitchen", 0.0043, "Switch to cast iron tawa/kadai or heavy tri-ply steel", 0.0008),
    "Particleboard / MDF Fast Furniture": ("Home & Kitchen", 0.0047, "Choose reclaimed teak hardwood or Cane/Rattan craft pieces", 0.0006),
    "Paraffin Wax Scented Candles": ("Home & Kitchen", 0.0036, "Switch to Phool temple-flower incense or soy/beeswax diyas", 0.0009),
    "Polyester Microfiber Blankets": ("Home & Kitchen", 0.0042, "Switch to Jaipur organic cotton razai or Desi wool blankets", 0.0010),
    "Incandescent / CFL Bulbs": ("Home & Kitchen", 0.0054, "Switch to 5-Star BEE-rated high-efficiency LED bulbs", 0.0008),
    "Single-Use Coffee Pods / Capsules": ("Home & Kitchen", 0.0044, "Use a traditional stainless steel coffee filter or French Press", 0.0005),
    "Bio-Enzyme / Refill Concentrates": ("Home & Kitchen", 0.0007, "Refill tablets cut water freight weight by 90%", 0.0006),
    "Stainless Steel Dabbas & Jars": ("Home & Kitchen", 0.0006, "Lifespan replaces thousands of single-use plastic containers", 0.0005),
    "Bamboo & Coconut Coir Home Goods": ("Home & Kitchen", 0.0008, "Bamboo products utilize fast-growing agricultural by-products", 0.0006),
    "Pre-Owned Solid Wood Furniture": ("Home & Kitchen", 0.0004, "Retains sequestered carbon inside durable timber", 0.0004),

    # --- 6. PERSONAL CARE & BEAUTY ---
    "Liquid Shampoo (Plastic Bottle)": ("Personal Care", 0.0034, "Switch to solid shampoo bars or shikakai/reetha", 0.0006),
    "Disposable Plastic Razors": ("Personal Care", 0.0038, "Switch to a brass/steel double-edge safety razor", 0.0005),
    "Liquid Body Wash (Plastic Pump)": ("Personal Care", 0.0033, "Switch to handmade soap bars or ubtan powder", 0.0007),
    "Aerosol Deodorant Spray (Gas)": ("Personal Care", 0.0049, "Switch to alum block (phitkari) or cream deodorant in tin", 0.0007),
    "Toothpaste in Laminate Tube": ("Personal Care", 0.0032, "Switch to toothpaste tablets in glass jars or datun", 0.0007),
    "Virgin Plastic Toothbrush": ("Personal Care", 0.0035, "Switch to compostable bamboo handle toothbrush", 0.0006),
    "Single-Use Cotton Pads & Wet Wipes": ("Personal Care", 0.0037, "Switch to washable organic cotton makeup rounds", 0.0005),
    "Cosmetics in Virgin Plastic Compacts": ("Personal Care", 0.0036, "Switch to refillable sticks or multi-use lip/cheek tins", 0.0009),
    "Disposable Sanitary Pads (Plastic)": ("Personal Care", 0.0034, "Switch to medical-grade silicone menstrual cup or cloth pads", 0.0004),
    "Chemical Sunscreen (Plastic Tube)": ("Personal Care", 0.0031, "Switch to non-nano zinc oxide mineral sunscreen in aluminum tins", 0.0008),
    "Zero-Waste Solid Soap & Shampoo": ("Personal Care", 0.0006, "Waterless bars eliminate plastic bottles and reduce shipping weight", 0.0005),

    # --- 7. TRAVEL, COMMUTE & MOBILITY ---
    "Short-Haul Domestic Flight (<1k km)": ("Travel & Commute", 0.0095, "Take Indian Railways electrified train (cuts CO2 by 80%)", 0.0012),
    "Long-Haul Economy Flight": ("Travel & Commute", 0.0088, "Fly direct routes (takeoff burns 25% of fuel) and pack light", 0.0045),
    "Business Class Flight Ticket": ("Travel & Commute", 0.0135, "Switch to Economy seating (Business footprint is 3x higher)", 0.0045),
    "Petrol / Diesel Fuel Fill-Up": ("Travel & Commute", 0.0082, "Use metro/suburban rail, carpool, or switch to an EV", 0.0014),
    "Solo Cab Ride (Ola / Uber Non-EV)": ("Travel & Commute", 0.0058, "Choose BluSmart (EV cabs), metro feeder, or shared transit", 0.0012),
    "Cruise Ship Holiday Package": ("Travel & Commute", 0.0092, "Swap for Konkan/mountain rail journeys or homestays", 0.0018),
    "Luxury Hotel Stay (24x7 AC)": ("Travel & Commute", 0.0039, "Stay at eco-homestays, set AC to 24-26°C, decline daily linen wash", 0.0013),
    "Indian Railways Intercity Ticket": ("Travel & Commute", 0.0009, "Electrified rail network with ultra-low per-passenger emissions", 0.0007),
    "City Metro / Bus Pass": ("Travel & Commute", 0.0006, "Mass rapid transit reduces urban traffic congestion", 0.0005),
    "Electric Vehicle (EV) Charging": ("Travel & Commute", 0.0008, "Daytime solar-hour charging minimizes lifecycle emissions", 0.0005),
    "Bicycle / E-Bike Purchase": ("Travel & Commute", 0.0004, "Zero tailpipe emissions; 40x lower lifecycle carbon than a petrol car", 0.0003),

    # --- 8. HOBBIES, PETS, OFFICE & GIFTS ---
    "Imported Meat-Heavy Pet Food": ("Hobbies & Pets", 0.0048, "Switch to freshly cooked rice/lentil meals or insect-protein pet food", 0.0013),
    "Strip-Mined Bentonite Cat Litter": ("Hobbies & Pets", 0.0039, "Switch to biodegradable pine wood pellets or walnut shell litter", 0.0008),
    "Brand New Hardcover Book (Virgin)": ("Hobbies & Pets", 0.0028, "Borrow from libraries, buy pre-owned, or use an e-reader", 0.0004),
    "Plastic-Coated / Foil Gift Wrapping": ("Hobbies & Pets", 0.0036, "Use reusable cotton potli bags or plantable seed-paper cards", 0.0005),
    "Hothouse Imported Cut Flower Bouquet": ("Hobbies & Pets", 0.0055, "Gift a potted indigenous houseplant in terracotta", 0.0008),
    "Virgin Plastic Battery Toys": ("Hobbies & Pets", 0.0044, "Switch to Channapatna natural lacquer wooden toys", 0.0009),
    "Peat-Moss Potting Mix & Urea": ("Hobbies & Pets", 0.0046, "Switch to coconut coir (cocopeat) + vermicompost", 0.0007),
    "Pre-Owned Books / Toy Swap": ("Hobbies & Pets", 0.0004, "Circular sharing eliminates virgin paper pulp and plastic demand", 0.0004),
}

SECTORS = sorted(list(set(val[0] for val in PRODUCT_CATALOG.values())))
SECTOR_TO_CATEGORIES = {
    sec: sorted([cat for cat, val in PRODUCT_CATALOG.items() if val[0] == sec])
    for sec in SECTORS
}

ECO_TIPS = [
    "Extending the active life of clothing by 9 months reduces its carbon footprint by 20–30%.",
    "Manufacturing a new smartphone accounts for ~80% of its lifetime CO2 emissions before first use.",
    "Indigenous millets (Ragi, Jowar) require 70% less water than flooded paddy rice.",
    "Indian Railways electrified intercity trains emit ~80% less CO2 per passenger-km than domestic flights.",
]


# ==============================================================================
# 3. CLIENT-SIDE URL COMPRESSION ENGINE (SURVIVES REFRESH & ZERO SYNC BUGS)
# ==============================================================================
def encode_purchases_to_url(purchases: list) -> str:
    """Compresses purchases into an ultra-compact Base64 URL string."""
    if not purchases:
        return ""
    minimal = [
        [
            p["date"],
            p["item"],
            p["category"],
            p["brand"],
            p["qty"],
            p["unit_price"],
            p["price"],
        ]
        for p in purchases
    ]
    raw = json.dumps(minimal).encode("utf-8")
    compressed = zlib.compress(raw, level=9)
    return base64.urlsafe_b64encode(compressed).decode("ascii")


def decode_purchases_from_url(encoded_str: str) -> list:
    """Restores the full analytical ledger from a compressed URL string upon refresh."""
    if not encoded_str:
        return []
    try:
        compressed = base64.urlsafe_b64decode(encoded_str.encode("ascii"))
        raw = zlib.decompress(compressed).decode("utf-8")
        minimal = json.loads(raw)
        reconstructed = []
        for d, item, cat, brand, qty, unit_p, tot_p in minimal:
            sec, mult, alt_desc, alt_mult = PRODUCT_CATALOG.get(
                cat, ("General", 0.001, "Sustainable alternative", 0.0005)
            )
            co2_val = round(tot_p * mult, 2)
            swap_val = round(tot_p * alt_mult, 2)
            reconstructed.append(
                {
                    "date": d,
                    "item": item,
                    "sector": sec,
                    "category": cat,
                    "brand": brand,
                    "qty": int(qty),
                    "unit_price": round(unit_p, 2),
                    "price": round(tot_p, 2),
                    "multiplier": mult,
                    "co2": co2_val,
                    "greener_swap": alt_desc,
                    "swap_co2": swap_val,
                    "potential_co2_saved": round(max(0.0, co2_val - swap_val), 2),
                }
            )
        return reconstructed
    except Exception:
        return []


# Load purchases from the active browser tab's URL param on refresh
if "purchases" not in st.session_state:
    url_data = st.query_params.get("ledger", "")
    st.session_state.purchases = decode_purchases_from_url(url_data)


# ==============================================================================
# 4. 10-STAGE AUTOMATIC TURTLE DRAWINGS
# ==============================================================================
class PurePythonTurtle:
    def __init__(self, width=480, height=225, bg_color="#EDE7DE"):
        self.scale = 3
        self.width = width * self.scale
        self.height = height * self.scale
        self.out_width = width
        self.out_height = height
        self.img = Image.new("RGB", (self.width, self.height), bg_color)
        self.draw = ImageDraw.Draw(self.img)
        self.x = self.width / 2.0
        self.y = self.height / 2.0
        self.angle = 0.0
        self.is_down = True
        self.is_filling = False
        self.pen_color = "#1B4332"
        self.fill_color = "#40916C"
        self.pen_width = 2 * self.scale
        self.poly_points = []
        self.stroke_queue = []

    def penup(self):
        self.is_down = False

    def pendown(self):
        self.is_down = True

    def pencolor(self, color: str):
        self.pen_color = color

    def fillcolor(self, color: str):
        self.fill_color = color

    def pensize(self, width: float):
        self.pen_width = max(1, int(width * self.scale))

    def setheading(self, angle: float):
        self.angle = float(angle)

    def left(self, angle: float):
        self.angle = (self.angle + angle) % 360

    def right(self, angle: float):
        self.angle = (self.angle - angle) % 360

    def forward(self, distance: float):
        rad = math.radians(self.angle)
        dist = distance * self.scale
        nx = self.x + dist * math.cos(rad)
        ny = self.y - dist * math.sin(rad)
        if self.is_down:
            if self.is_filling:
                self.stroke_queue.append(((self.x, self.y, nx, ny), self.pen_color, self.pen_width))
            else:
                self.draw.line([(self.x, self.y), (nx, ny)], fill=self.pen_color, width=self.pen_width)
        if self.is_filling:
            self.poly_points.append((nx, ny))
        self.x, self.y = nx, ny

    def goto(self, x: float, y: float):
        sx = self.width / 2.0 + (x * self.scale)
        sy = self.height / 2.0 - (y * self.scale)
        if self.is_down:
            if self.is_filling:
                self.stroke_queue.append(((self.x, self.y, sx, sy), self.pen_color, self.pen_width))
            else:
                self.draw.line([(self.x, self.y), (sx, sy)], fill=self.pen_color, width=self.pen_width)
        if self.is_filling:
            self.poly_points.append((sx, sy))
        self.x, self.y = sx, sy

    def begin_fill(self):
        self.is_filling = True
        self.poly_points = [(self.x, self.y)]
        self.stroke_queue = []

    def end_fill(self):
        if self.is_filling and len(self.poly_points) > 2:
            self.draw.polygon(self.poly_points, fill=self.fill_color)
            for coords, col, w in self.stroke_queue:
                self.draw.line([(coords[0], coords[1]), (coords[2], coords[3])], fill=col, width=w)
        self.is_filling = False
        self.poly_points = []
        self.stroke_queue = []

    def circle(self, radius: float, extent: float = 360.0, steps: int = 60):
        step_len = (2.0 * math.pi * radius * (extent / 360.0)) / steps
        step_ang = extent / steps
        for _ in range(steps):
            self.forward(step_len)
            self.left(step_ang)

    def get_image(self) -> Image.Image:
        return self.img.resize((self.out_width, self.out_height), Image.Resampling.LANCZOS)


def get_automatic_badge_stage(total_co2: float, monthly_budget: float) -> tuple:
    pct = (total_co2 / monthly_budget * 100.0) if monthly_budget > 0 else 0.0
    rem = max(0.0, monthly_budget - total_co2)

    if total_co2 <= 0.0:
        return ("sunrise", "Fresh Start", "0 kg CO₂e emitted so far. Log your first purchase on the left to start tracking.", "1%–12% budget: Tiny Sprout")
    if pct <= 12.0:
        return ("sprout", "Tiny Sprout", f"Only {pct:.0f}% of your carbon budget used ({total_co2:.1f} kg). Your footprint is minimal.", f"Keep under {monthly_budget*0.25:.1f} kg for Blooming Flower")
    if pct <= 25.0:
        return ("flower", "Blooming Eco Flower", f"Very low emissions — {pct:.0f}% of budget used with {rem:.1f} kg still available.", f"Next at 25%: Lush Pine Tree")
    if pct <= 40.0:
        return ("pine", "Lush Pine Tree", f"Low, nature-friendly footprint — {pct:.0f}% of budget used ({rem:.1f} kg left).", f"Next at 40%: Gold Eco Saver Trophy")
    if pct <= 55.0:
        return ("trophy", "Gold Eco Saver Trophy", f"Halfway mark ({pct:.0f}% used) — you are well within a sustainable monthly pace.", f"Next at 55%: Green Guardian Shield")
    if pct <= 70.0:
        return ("shield", "Green Guardian Shield", f"Good standing — {pct:.0f}% of budget used ({rem:.1f} kg remaining before target limit).", f"Next at 70%: Botanical Leaf")
    if pct <= 82.0:
        return ("leaf", "Botanical Leaf", f"Moderate emissions ({pct:.0f}% used) — choose greener swaps for upcoming purchases.", f"Next at 82%: Clean Energy Windmill")
    if pct <= 92.0:
        return ("windmill", "Clean Energy Windmill", f"Approaching monthly limit ({pct:.0f}% used, {rem:.1f} kg left) — prioritize low-impact items.", f"Next at 92%: Amber Hourglass Warning")
    if pct <= 100.0:
        return ("hourglass", "Amber Hourglass Warning", f"Almost out of carbon budget ({pct:.0f}% used) — only {rem:.1f} kg of allowance left.", f"Over 100%: Heavy Carbon Footprint")
    return ("footprint", "Heavy Carbon Footprint", f"Monthly carbon budget exceeded by {total_co2 - monthly_budget:.1f} kg ({pct:.0f}% of target). Check Compare Alternatives to cut emissions.", "Reduce high-impact purchases")


def generate_automatic_turtle_image(stage_key: str, utilization: float) -> Image.Image:
    t = PurePythonTurtle(width=480, height=220, bg_color="#EDE7DE")

    for radius, col, w in [(86, "#D4CBBF", 4.5), (72, "#DFD7CA", 1.2)]:
        t.penup()
        t.goto(0, -radius)
        t.setheading(0)
        t.pendown()
        t.pencolor(col)
        t.pensize(w)
        t.circle(radius)

    sweep_deg = min(360.0, max(0.0, utilization * 360.0))
    if sweep_deg > 1.0:
        arc_col = "#2D6A4F" if utilization <= 0.55 else ("#52B788" if utilization <= 0.82 else ("#D97706" if utilization <= 1.00 else "#DC2626"))
        t.penup()
        t.goto(0, 86)
        t.setheading(0)
        t.pendown()
        t.pencolor(arc_col)
        t.pensize(5.0)
        steps = max(6, int(sweep_deg / 5))
        step_len = (2.0 * math.pi * 86.0 * (sweep_deg / 360.0)) / steps
        step_ang = sweep_deg / steps
        for _ in range(steps):
            t.forward(step_len)
            t.right(step_ang)

    if stage_key == "sunrise":
        t.penup(); t.goto(-48, -18); t.pendown(); t.pencolor("#1B4332"); t.fillcolor("#40916C"); t.pensize(2); t.begin_fill()
        t.goto(48, -18); t.goto(36, -34); t.goto(-36, -34); t.goto(-48, -18); t.end_fill()
        t.penup(); t.goto(28, -18); t.setheading(90); t.pendown(); t.pencolor("#B45309"); t.fillcolor("#FBBF24"); t.begin_fill()
        t.circle(28, 180, steps=30); t.goto(28, -18); t.end_fill()
        for ang in [30, 60, 90, 120, 150]:
            rad = math.radians(ang)
            t.penup(); t.goto(34 * math.cos(rad), -18 + 34 * math.sin(rad)); t.pendown(); t.pencolor("#D97706"); t.pensize(2.2)
            t.goto(46 * math.cos(rad), -18 + 46 * math.sin(rad))

    elif stage_key == "sprout":
        t.penup(); t.goto(-34, -36); t.pendown(); t.pencolor("#5C4033"); t.fillcolor("#8D6E63"); t.pensize(2); t.begin_fill()
        t.goto(34, -36); t.goto(22, -22); t.goto(-22, -22); t.goto(-34, -36); t.end_fill()
        t.penup(); t.goto(0, -22); t.pendown(); t.pencolor("#1B4332"); t.pensize(3.2); t.goto(0, 18)
        for sign in [-1, 1]:
            t.penup(); t.goto(0, 2); t.pendown(); t.pencolor("#182E23"); t.fillcolor("#52B788"); t.pensize(1.8); t.begin_fill()
            t.goto(sign * 18, 10); t.goto(sign * 28, 26); t.goto(sign * 10, 22); t.goto(0, 10); t.goto(0, 2); t.end_fill()

    elif stage_key == "flower":
        for ang in range(0, 360, 45):
            rad = math.radians(ang)
            px, py = 24 * math.cos(rad), 24 * math.sin(rad)
            t.penup(); t.goto(px, py - 13); t.setheading(0); t.pendown(); t.pencolor("#1B4332"); t.fillcolor("#74C69D"); t.pensize(1.6); t.begin_fill()
            t.circle(13, 360, steps=24); t.end_fill()
        t.penup(); t.goto(0, -14); t.setheading(0); t.pendown(); t.pencolor("#9A6B26"); t.fillcolor("#FBBF24"); t.pensize(2); t.begin_fill()
        t.circle(14, 360, steps=28); t.end_fill()

    elif stage_key == "pine":
        t.penup(); t.goto(-7, -46); t.pendown(); t.pencolor("#5C4033"); t.fillcolor("#7F5539"); t.pensize(1.8); t.begin_fill()
        t.goto(7, -46); t.goto(7, -22); t.goto(-7, -22); t.goto(-7, -46); t.end_fill()
        tiers = [(-38, -22, 38, -22, 0, 6, "#1B4332"), (-31, -4, 31, -4, 0, 24, "#2D6A4F"), (-23, 14, 23, 14, 0, 44, "#40916C")]
        for x1, y1, x2, y2, top_x, top_y, col in tiers:
            t.penup(); t.goto(x1, y1); t.pendown(); t.pencolor("#182E23"); t.fillcolor(col); t.pensize(1.8); t.begin_fill()
            t.goto(x2, y2); t.goto(top_x, top_y); t.goto(x1, y1); t.end_fill()

    elif stage_key == "trophy":
        t.penup(); t.goto(-22, -42); t.pendown(); t.pencolor("#6E4711"); t.fillcolor("#9A6B26"); t.pensize(2); t.begin_fill()
        t.goto(22, -42); t.goto(18, -32); t.goto(-18, -32); t.goto(-22, -42); t.end_fill()
        t.penup(); t.goto(-5, -32); t.pendown(); t.fillcolor("#D4A373"); t.begin_fill()
        t.goto(5, -32); t.goto(5, -16); t.goto(-5, -16); t.goto(-5, -32); t.end_fill()
        for sign in [-1, 1]:
            t.penup(); t.goto(sign * 22, 22); t.pendown(); t.pencolor("#8C5E1A"); t.pensize(2.5)
            t.goto(sign * 34, 18); t.goto(sign * 30, 0); t.goto(sign * 16, -4)
        t.penup(); t.goto(-26, 28); t.pendown(); t.pencolor("#6E4711"); t.fillcolor("#FBBF24"); t.pensize(2); t.begin_fill()
        t.goto(26, 28); t.goto(16, -16); t.goto(-16, -16); t.goto(-26, 28); t.end_fill()

    elif stage_key == "shield":
        t.penup(); t.goto(0, 44); t.pendown(); t.pencolor("#182E23"); t.fillcolor("#2D6A4F"); t.pensize(2.2); t.begin_fill()
        t.goto(34, 28); t.goto(28, -12); t.goto(0, -44); t.goto(-28, -12); t.goto(-34, 28); t.goto(0, 44); t.end_fill()
        t.penup(); t.goto(-14, -2); t.pendown(); t.pencolor("#FFFFFF"); t.pensize(4.0); t.goto(-3, -14); t.goto(16, 12)

    elif stage_key == "leaf":
        t.penup(); t.goto(0, -40); t.pendown(); t.pencolor("#182E23"); t.fillcolor("#52B788"); t.pensize(2.0); t.begin_fill()
        t.goto(24, -18); t.goto(28, 6); t.goto(15, 28); t.goto(0, 44); t.goto(-15, 28); t.goto(-28, 6); t.goto(-24, -18); t.goto(0, -40); t.end_fill()
        t.penup(); t.goto(0, -48); t.pendown(); t.pencolor("#182E23"); t.pensize(2.2); t.goto(0, 34)
        for vy, span in [(-16, 15), (0, 18), (16, 11)]:
            for sign in [-1, 1]:
                t.penup(); t.goto(0, vy); t.pendown(); t.pensize(1.5); t.goto(sign * span, vy + 9)

    elif stage_key == "windmill":
        t.penup(); t.goto(-7, -46); t.pendown(); t.pencolor("#1B4332"); t.fillcolor("#D5CFC0"); t.pensize(2.0); t.begin_fill()
        t.goto(7, -46); t.goto(2.5, 8); t.goto(-2.5, 8); t.goto(-7, -46); t.end_fill()
        for ang in [90, 210, 330]:
            rad = math.radians(ang); perp = math.radians(ang + 90)
            bx, by = 40 * math.cos(rad), 8 + 40 * math.sin(rad)
            wx, wy = 5 * math.cos(perp), 5 * math.sin(perp)
            t.penup(); t.goto(0, 8); t.pendown(); t.pencolor("#1B4332"); t.fillcolor("#40916C"); t.pensize(1.8); t.begin_fill()
            t.goto(wx, 8 + wy); t.goto(bx, by); t.goto(-wx, 8 - wy); t.goto(0, 8); t.end_fill()
        t.penup(); t.goto(0, 2); t.setheading(0); t.pendown(); t.fillcolor("#182E23"); t.begin_fill(); t.circle(6, 360, steps=18); t.end_fill()

    elif stage_key == "hourglass":
        t.penup(); t.goto(-24, 36); t.pendown(); t.pencolor("#92400E"); t.fillcolor("#FEF3C7"); t.pensize(2.2); t.begin_fill()
        t.goto(24, 36); t.goto(0, 0); t.goto(-24, 36); t.end_fill()
        t.penup(); t.goto(-24, -36); t.pendown(); t.fillcolor("#F59E0B"); t.begin_fill()
        t.goto(24, -36); t.goto(0, 0); t.goto(-24, -36); t.end_fill()
        for fy in [38, -38]:
            t.penup(); t.goto(-28, fy); t.pendown(); t.pencolor("#78350F"); t.pensize(4.0); t.goto(28, fy)

    else:
        t.penup(); t.goto(0, -56); t.setheading(0); t.pendown(); t.pencolor("#EF4444"); t.pensize(1.8); t.circle(56)
        t.penup(); t.goto(0, -38); t.setheading(0); t.pendown(); t.pencolor("#991B1B"); t.fillcolor("#DC2626"); t.pensize(1.8); t.begin_fill()
        t.circle(20); t.end_fill()
        t.penup(); t.goto(0, 8); t.pendown(); t.begin_fill(); t.circle(12); t.end_fill()
        for idx, x_off in enumerate([-15, -7.5, 0, 7.5, 15]):
            t.penup(); t.goto(x_off, 33 + (4 - abs(idx - 2) * 2)); t.pendown(); t.begin_fill(); t.circle(4.0 - abs(idx - 2) * 0.45); t.end_fill()

    return t.get_image()


# ==============================================================================
# 5. EXCEL (.XLSX) & PDF EXPORT GENERATORS
# ==============================================================================
def build_autofit_xlsx(df_export: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_export.to_excel(writer, index=False, sheet_name="ShopImpact Log")
        worksheet = writer.sheets["ShopImpact Log"]
        for idx, col in enumerate(df_export.columns):
            max_len = max(df_export[col].astype(str).map(len).max(), len(str(col))) + 4
            worksheet.set_column(idx, idx, max_len)
    return output.getvalue()


def sanitize_text(text: str) -> str:
    return str(text).encode("latin-1", "ignore").decode("latin-1").strip()


def generate_pdf_report(df: pd.DataFrame, total_spend: float, total_co2: float, total_saved: float, badge_name: str) -> bytes:
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    safe_badge = sanitize_text(badge_name)

    pdf.set_fill_color(27, 67, 50)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 15)
    pdf.cell(0, 13, "  ShopImpact | Monthly Carbon Footprint Report", ln=True, fill=True)
    pdf.ln(4)

    pdf.set_text_color(24, 46, 35)
    pdf.set_font("Helvetica", "B", 10.5)
    pdf.cell(0, 7, f"Date: {date.today().strftime('%d %b %Y')}   |   Current Standing: {safe_badge}", ln=True)
    pdf.set_font("Helvetica", "", 9.5)
    pdf.cell(
        0,
        6,
        f"Total Spend: Rs. {total_spend:,.2f}   |   Estimated Footprint: {total_co2:,.2f} kg CO2e   |   Potential Savings: -{total_saved:,.2f} kg CO2e",
        ln=True,
    )
    pdf.ln(4)

    pdf.set_fill_color(220, 215, 201)
    pdf.set_text_color(24, 46, 35)
    pdf.set_font("Helvetica", "B", 8)
    pdf.cell(22, 8, "Date", border=1, fill=True)
    pdf.cell(34, 8, "Item", border=1, fill=True)
    pdf.cell(48, 8, "Category", border=1, fill=True)
    pdf.cell(14, 8, "Qty", border=1, fill=True, align="C")
    pdf.cell(26, 8, "Total (INR)", border=1, fill=True, align="R")
    pdf.cell(23, 8, "CO2e (kg)", border=1, fill=True, align="R")
    pdf.cell(23, 8, "Saveable", border=1, ln=True, fill=True, align="R")

    pdf.set_font("Helvetica", "", 8)
    for _, row in df.iterrows():
        pdf.cell(22, 7, sanitize_text(row["date"]), border=1)
        pdf.cell(34, 7, sanitize_text(row["item"])[:19], border=1)
        pdf.cell(48, 7, sanitize_text(row["category"])[:28], border=1)
        pdf.cell(14, 7, str(int(row.get("qty", 1))), border=1, align="C")
        pdf.cell(26, 7, f"Rs. {row['price']:,.0f}", border=1, align="R")
        pdf.cell(23, 7, f"{row['co2']:,.2f} kg", border=1, align="R")
        pdf.cell(23, 7, f"-{row['potential_co2_saved']:,.2f} kg", border=1, ln=True, align="R")

    pdf.ln(5)
    pdf.set_font("Helvetica", "I", 8)
    pdf.set_text_color(60, 60, 60)
    pdf.multi_cell(
        0,
        4.5,
        "Note: Carbon multipliers (kg CO2e per INR) are calibrated from lifecycle emission inventories. "
        "Recommended alternatives highlight practical lower-carbon choices.",
    )
    return bytes(pdf.output())


# ==============================================================================
# 6. SIDEBAR: PURCHASE ENTRY & BUDGET TARGET
# ==============================================================================
with st.sidebar:
    st.subheader("ShopImpact")
    st.markdown(f"Reference Catalog: **{len(PRODUCT_CATALOG)} Categories**")

    monthly_budget = st.slider("Monthly Carbon Target (kg CO₂e)", min_value=15, max_value=250, value=60, step=5)

    st.divider()
    st.subheader("Record a Purchase")

    selected_sector = st.selectbox("Sector", SECTORS)
    selected_cat = st.selectbox("Category", SECTOR_TO_CATEGORIES[selected_sector])

    sec_info, preview_mult, preview_alt, preview_alt_mult = PRODUCT_CATALOG[selected_cat]
    per_1000_kg = preview_mult * 1000

    with st.container(border=True):
        st.markdown(f"**Emission Intensity:** `{per_1000_kg:.2f} kg CO₂e / ₹1,000`")
        st.markdown(f"**Greener Alternative:** {preview_alt}")

    item_input = st.text_input("Item Description", placeholder="e.g., Cotton Kurta / Running Shoes", key="inp_item")
    brand_input = st.text_input("Brand / Store", placeholder="e.g., Fabindia / Local Store", key="inp_brand")

    p_col1, p_col2 = st.columns([1.3, 1])
    with p_col1:
        unit_price_input = st.number_input("Unit Price (₹)", min_value=10.0, max_value=500000.0, value=1200.0, step=100.0, key="inp_price")
    with p_col2:
        qty_input = st.selectbox("Quantity", options=list(range(1, 201)), index=0, key="inp_qty")

    date_input = st.date_input("Purchase Date", value=date.today(), key="inp_date")

    total_line_price = round(unit_price_input * qty_input, 2)
    est_line_co2 = round(total_line_price * preview_mult, 2)
    st.markdown(f"**Total:** ₹{total_line_price:,.0f} ({qty_input} × ₹{unit_price_input:,.0f})  •  **Est. CO₂:** {est_line_co2:.2f} kg")

    if st.button("Add to Monthly Log", type="primary", use_container_width=True):
        swap_val = round(total_line_price * preview_alt_mult, 2)
        formatted_date = date_input.strftime("%d %b %Y")
        st.session_state.purchases.append(
            {
                "date": formatted_date,
                "item": item_input.strip() if item_input.strip() else selected_cat.split("(")[0].strip(),
                "sector": selected_sector,
                "category": selected_cat,
                "brand": brand_input.strip() if brand_input.strip() else "General Store",
                "qty": int(qty_input),
                "unit_price": round(unit_price_input, 2),
                "price": total_line_price,
                "multiplier": preview_mult,
                "co2": est_line_co2,
                "greener_swap": preview_alt,
                "swap_co2": swap_val,
                "potential_co2_saved": round(max(0.0, est_line_co2 - swap_val), 2),
            }
        )
        # Update client-side URL so refresh preserves the data
        st.query_params["ledger"] = encode_purchases_to_url(st.session_state.purchases)
        st.rerun()

# ==============================================================================
# 7. TOP HEADER BAR & KPI METRICS
# ==============================================================================
head_l, head_r = st.columns([3, 1])
with head_l:
    st.title("ShopImpact")
    st.markdown("Conscious Shopping & Monthly Carbon Footprint Tracker (INR)")
with head_r:
    st.write("")
    st.markdown(f"**Monthly Quota:** `{monthly_budget} kg CO₂e`")

st.divider()

df = pd.DataFrame(st.session_state.purchases)
total_spend = df["price"].sum() if not df.empty else 0.0
total_units = int(df["qty"].sum()) if (not df.empty and "qty" in df.columns) else len(df)
total_co2 = df["co2"].sum() if not df.empty else 0.0
total_potential_save = df["potential_co2_saved"].sum() if not df.empty else 0.0
avg_per_1000 = ((total_co2 / total_spend) * 1000) if total_spend > 0 else 0.0
low_impact_count = int((df["multiplier"] <= 0.0015).sum()) if not df.empty else 0
utilization = (total_co2 / monthly_budget) if monthly_budget > 0 else 0.0

stage_key, badge_title, badge_desc, next_stage_hint = get_automatic_badge_stage(total_co2, monthly_budget)

k1, k2, k3, k4 = st.columns(4)
with k1:
    with st.container(border=True):
        st.metric("Monthly Spend", f"₹{total_spend:,.0f}")
        st.markdown(f"**{len(df)}** purchases (**{total_units}** items)")
with k2:
    rem = monthly_budget - total_co2
    rem_text = f"**{rem:,.1f} kg** safe remaining" if rem >= 0 else f"**{abs(rem):,.1f} kg** over quota"
    with st.container(border=True):
        st.metric("Estimated CO₂ Footprint", f"{total_co2:,.2f} kg")
        st.markdown(rem_text)
with k3:
    with st.container(border=True):
        st.metric("Carbon Intensity", f"{avg_per_1000:.2f} kg / ₹1k")
        st.markdown(f"**{low_impact_count}** low-impact choices")
with k4:
    with st.container(border=True):
        st.metric("Avoidable Emissions", f"{total_potential_save:,.2f} kg")
        st.markdown("Saveable via brand swaps")

bar_ratio = min(1.0, max(0.0, utilization))
st.progress(bar_ratio, text=f"Carbon Budget Utilization: {total_co2:.1f} / {monthly_budget:.1f} kg CO₂e ({utilization*100:.0f}%)")

# ==============================================================================
# 8. MAIN WORKSPACE TABS
# ==============================================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "Dashboard",
    "Compare Alternatives",
    "Category Directory (108)",
    "Purchase Log & Export",
])

with tab1:
    col_l, col_r = st.columns([1.55, 1.2])
    with col_l:
        with st.container(border=True):
            st.subheader("Emissions by Product Category")
            st.markdown("Cumulative carbon footprint (kg CO₂e) across your logged purchases")
            if not df.empty:
                sec_df = df.groupby("category", as_index=False).agg(total_co2=("co2", "sum"), spend=("price", "sum"), units=("qty", "sum"))
                
                chart = (
                    alt.Chart(sec_df)
                    .mark_bar(cornerRadiusTopRight=5, cornerRadiusBottomRight=5)
                    .encode(
                        y=alt.Y("category:N", sort="-x", title=None, axis=alt.Axis(labelLimit=260, labelFontSize=12)),
                        x=alt.X("total_co2:Q", title="Emissions (kg CO₂e)", axis=alt.Axis(labelFontSize=12)),
                        color=alt.Color("total_co2:Q", scale=alt.Scale(scheme="greens"), legend=None),
                        tooltip=["category", "units", "spend", "total_co2"],
                    )
                    .properties(height=alt.Step(38))
                )
                st.altair_chart(chart, use_container_width=True)
            else:
                st.markdown("No purchases logged yet. Use the sidebar form to add your first entry.")

        with st.container(border=True):
            st.markdown("💡 **Sustainability Insight**")
            st.markdown(random.choice(ECO_TIPS))

    with col_r:
        with st.container(border=True):
            st.subheader(f"{badge_title}")
            st.markdown(badge_desc)
            badge_img = generate_automatic_turtle_image(stage_key=stage_key, utilization=utilization)
            st.image(badge_img, use_container_width=True)
            st.markdown(f"**Carbon Gauge:** `{utilization*100:.0f}%` of {monthly_budget} kg allowance  •  **Target:** {next_stage_hint}")

with tab2:
    sim_c1, sim_c2 = st.columns([1.15, 1.4])
    with sim_c1:
        with st.container(border=True):
            st.subheader("Pre-Purchase Comparison")
            st.markdown("Evaluate potential carbon savings before committing to a purchase.")
            s_sec = st.selectbox("Sector", SECTORS, key="sim_sec")
            s_cat = st.selectbox("Product Category", SECTOR_TO_CATEGORIES[s_sec], key="sim_cat")
            sc1, sc2 = st.columns([1.4, 1])
            with sc1:
                s_unit_pr = st.number_input("Unit Price (₹ INR)", min_value=50.0, max_value=200000.0, value=1500.0, step=100.0, key="sim_pr")
            with sc2:
                s_qty = st.selectbox("Quantity", options=list(range(1, 201)), index=0, key="sim_qty")

    s_total_pr = round(s_unit_pr * s_qty, 2)
    base_mult = PRODUCT_CATALOG[s_cat][1]
    alt_desc = PRODUCT_CATALOG[s_cat][2]
    alt_mult = PRODUCT_CATALOG[s_cat][3]
    planned_co2 = round(s_total_pr * base_mult, 2)
    swapped_co2 = round(s_total_pr * alt_mult, 2)
    saved_co2 = round(max(0.0, planned_co2 - swapped_co2), 2)
    pct_cut = round((saved_co2 / planned_co2) * 100, 0) if planned_co2 > 0 else 0

    with sim_c2:
        with st.container(border=True):
            st.subheader(f"Impact: {s_qty} × ₹{s_unit_pr:,.0f} (₹{s_total_pr:,.0f})")
            st.markdown(f"**Category:** {s_cat}")
            m_col1, m_col2 = st.columns(2)
            with m_col1:
                with st.container(border=True):
                    st.metric("Standard Choice", f"{planned_co2:.2f} kg CO₂e", "Higher footprint", delta_color="inverse")
                    st.caption(f"Factor: {base_mult*1000:.2f} kg / ₹1k")
            with m_col2:
                with st.container(border=True):
                    st.metric("Greener Swap", f"{swapped_co2:.2f} kg CO₂e", f"-{saved_co2:.2f} kg ({pct_cut:.0f}% lower)", delta_color="normal")
                    st.caption(f"Factor: {alt_mult*1000:.2f} kg / ₹1k")
            st.markdown(f"**Recommended Alternative:** {alt_desc}")

    st.divider()
    st.subheader("Alternatives for Your Logged Purchases")
    st.markdown("Retrospective recommendations for items already recorded in your monthly ledger.")

    if not df.empty:
        for idx, row in df[::-1].iterrows():
            q_val = int(row.get("qty", 1))
            u_val = row.get("unit_price", row["price"])
            with st.container(border=True):
                r_left, r_right = st.columns([3, 1])
                with r_left:
                    st.markdown(f"**{row['item']}** — Qty: {q_val} × ₹{u_val:,.0f} (`₹{row['price']:,.0f}`)")
                    st.caption(f"{row['category']}  •  {row['brand']}  •  {row['date']}")
                with r_right:
                    st.markdown(f"**{row['co2']} kg CO₂e**")
                    st.caption(f"Saveable: -{row['potential_co2_saved']} kg")
                st.markdown(f"**Recommended Swap:** {row['greener_swap']} *(reduces footprint to `{row['swap_co2']} kg CO₂e`)*")
    else:
        st.markdown("No purchases recorded yet. Add an entry from the sidebar to view tailored swap recommendations.")

with tab3:
    q_filter = st.text_input("Search catalog by category, sector, or sustainable alternative", placeholder="e.g., millets, refurbished, bamboo, rail, cotton...")
    ref_df = pd.DataFrame(
        [
            {
                "Sector": v[0],
                "Category": k,
                "kg CO₂e / ₹1,000": round(v[1] * 1000, 2),
                "Swap kg CO₂e / ₹1,000": round(v[3] * 1000, 2),
                "Recommended Greener Alternative": v[2],
            }
            for k, v in PRODUCT_CATALOG.items()
        ]
    )
    if q_filter.strip():
        mask = (
            ref_df["Category"].str.contains(q_filter, case=False, na=False)
            | ref_df["Recommended Greener Alternative"].str.contains(q_filter, case=False, na=False)
            | ref_df["Sector"].str.contains(q_filter, case=False, na=False)
        )
        ref_df = ref_df[mask]

    st.dataframe(
        ref_df,
        use_container_width=True,
        hide_index=True,
        height=440,
        column_config={
            "Category": st.column_config.TextColumn("Category", width="medium"),
            "Recommended Greener Alternative": st.column_config.TextColumn("Recommended Greener Alternative", width="large"),
        },
    )

with tab4:
    if not df.empty:
        if "qty" not in df.columns:
            df["qty"] = 1
        if "unit_price" not in df.columns:
            df["unit_price"] = df["price"]

        df["qty"] = df["qty"].astype(int)

        df_display = df[
            ["date", "item", "brand", "category", "qty", "unit_price", "price", "co2", "greener_swap", "potential_co2_saved"]
        ].rename(
            columns={
                "date": "Date",
                "item": "Item",
                "brand": "Brand / Store",
                "category": "Category",
                "qty": "Qty",
                "unit_price": "Unit Price (₹)",
                "price": "Total Price (₹)",
                "co2": "CO₂e (kg)",
                "greener_swap": "Recommended Alternative",
                "potential_co2_saved": "Saveable CO₂e (kg)",
            }
        )

        st.dataframe(
            df_display,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Category": st.column_config.TextColumn("Category", width="medium"),
                "Recommended Alternative": st.column_config.TextColumn("Recommended Alternative", width="large"),
            },
        )

        st.divider()
        d1, d2, d3, d4 = st.columns(4)

        with d1:
            xlsx_bytes = build_autofit_xlsx(df_display)
            st.download_button(
                "Download Excel (.xlsx)",
                xlsx_bytes,
                f"ShopImpact_Report_{date.today()}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

        with d2:
            st.download_button(
                "Download CSV (.csv)",
                df_display.to_csv(index=False).encode("utf-8-sig"),
                f"ShopImpact_Log_{date.today()}.csv",
                "text/csv",
                use_container_width=True,
            )

        with d3:
            if PDF_AVAILABLE:
                st.download_button(
                    "Download PDF (.pdf)",
                    generate_pdf_report(df, total_spend, total_co2, total_potential_save, badge_title),
                    f"ShopImpact_Report_{date.today()}.pdf",
                    "application/pdf",
                    use_container_width=True,
                )
            else:
                st.markdown("Install `fpdf2` for PDF export.")

        with d4:
            if st.button("Clear Ledger", use_container_width=True):
                st.session_state.purchases = []
                if "ledger" in st.query_params:
                    del st.query_params["ledger"]
                st.rerun()
    else:
        st.markdown("Your purchase ledger is currently empty. Record a transaction from the sidebar to view and export logs.")
