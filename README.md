# 🌿 ShopImpact: Conscious Shopping Dashboard

[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://idai102-1000679-gaurav-bardia.streamlit.app)
[![Python](https://img.shields.io/badge/Python-3.11%2B-3B82F6?style=for-the-badge&logo=python&logoColor=white)](https://idai102-1000679-gaurav-bardia.streamlit.app)
[![Streamlit](https://img.shields.io/badge/Streamlit-Web_App-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)](https://idai102-1000679-gaurav-bardia.streamlit.app)
[![Status](https://img.shields.io/badge/Status-Production_Ready-22C55E?style=for-the-badge)](https://idai102-1000679-gaurav-bardia.streamlit.app)

> **🌐 Live Web Application:** [**Click Here to Launch ShopImpact on Streamlit Cloud**](https://idai102-1000679--gaurav-bardia.streamlit.app/)  

---

## 📖 Project Overview
**ShopImpact** is an interactive, Python-based conscious consumerism application built with **Streamlit**. It transforms everyday shopping routines into a mindful, eco-conscious experience. By leveraging curated spend-based carbon emission multipliers (EPA/DEFRA standards), the app translates the financial cost of purchases (in INR ₹) into an estimated environmental cost (kg CO₂e). It gently nudges users toward sustainable choices through an interactive "What-If" simulator, automated green alternatives, and a dynamically drawn 10-stage Turtle graphics reward system.

---

## 🎯 The Problem It Solves
1. **The Invisible Footprint:** Consumers usually cannot correlate a currency amount (e.g., ₹1,500) with its hidden environmental burden. ShopImpact makes this invisible cost immediately visible.
2. **Eco-Guilt vs. Positive Action:** Most carbon trackers lecture users, causing anxiety. ShopImpact uses positive reinforcement and gamified achievements (badges and blooming trees) to empower users.
3. **Information Asymmetry:** Shoppers often want to buy sustainably but don't know the exact alternatives. This app maps 108 product categories directly to actionable, specific brand and material swaps.
4. **Data Privacy:** Traditional tracking apps harvest user financial data. ShopImpact relies on isolated, pure in-memory `st.session_state` logic, keeping user data strictly private per device.

**How ShopImpact Solves This:**  
The app provides a real-time, side-by-side comparison of standard purchases against greener alternatives, allowing users to tangibly see the exact amount of CO₂e they can save *before* making a purchase.

---

## 👥 Key Stakeholders & Target Audience

| Stakeholder Group | Role / Interest | How ShopImpact Benefits Them |
| :--- | :--- | :--- |
| **🎓 Young Adults & Students** | Primary End Users | Fast, mobile-friendly interface to track fast-fashion and electronics footprints, gamified with visual badges. |
| **👨‍👩‍👧 Eco-Conscious Families** | Secondary Users | Tracks household grocery and utility footprints against a defined monthly carbon quota. |
| **🏢 SME Procurement Teams** | Commercial Users | Utilizes the 108-category database and Auto-Fit Excel / PDF ledger exports to audit and reduce office supply emissions. |

---

## ✨ Complete Feature Set

### 1. 🛒 Conscious Purchase Logging & Live Analytics
* **108-Category Database:** Covers 8 major sectors (Apparel, Food, Tech, Commute, etc.) with precise kg CO₂e / ₹1,000 multipliers.
* **Monthly Dashboard:** Real-time KPI metric cards tracking Total Spend (₹), Estimated Footprint (kg CO₂e), Carbon Intensity, and Potential Savings.
* **Dynamic Budget Progress:** A live progress bar mapping total emissions against a user-defined monthly carbon allowance.

### 2. 🐢 10-Stage Reactive Turtle Graphics Engine (Pure Python)
* Built entirely without HTML/JS, utilizing Python's `PIL` (Pillow) library to simulate Turtle vector drawing.
* The artwork dynamically evolves across **10 distinct stages** based on the user's carbon budget utilization:
  * *0% Used:* **Fresh Start** (Rising sun)
  * *1% – 25% Used:* **Tiny Sprout** to **Blooming Flower**
  * *25% – 70% Used:* **Lush Pine Tree** to **Green Guardian Shield**
  * *92% – 100% Used:* **Amber Hourglass Warning**
  * *>100% Used:* **Heavy Carbon Footprint** (Red alert)
* Features a live 360-degree carbon gauge arc that physically updates with every logged kilogram of CO₂e.

### 3. 🔄 Pre-Purchase 'What-If' Simulator
* A dedicated tab allowing users to test a purchase price and category *before* buying.
* Displays a side-by-side metric comparison: **Standard Option vs. Greener Alternative**.
* Calculates exact CO₂e avoided and displays actionable, verified swaps (e.g., *Switching to Fairphone, Thrift stores, or Oat Milk*).

### 4. 📥 Zero-Dependency Export Engine
* **Auto-Fit Excel (.xlsx):** Uses Python's `zipfile` and `xlsxwriter` engines to mathematically auto-size columns for perfectly formatted spreadsheet ledgers.
* **Professional PDF Audit Report:** Built with `fpdf2`, generating a sleek, printable document of the user's monthly climate impact, current standing, and detailed purchase rows.
* **CSV Logging:** Standard comma-separated export using `utf-8-sig` to ensure precise INR (₹) symbol rendering in Excel.

---

## 🎨 User Interface (UI) Design
* **100% Native Streamlit Engine:** Achieves a bespoke aesthetic utilizing `.streamlit/config.toml` without relying on messy CSS injections.
* **Adaptive Light/Dark Mode:** 
  * *Light Theme:* Warm organic sand (`#F6F5F0`) with deep forest green accents (`#1B4332`), evoking natural linen and earth.
  * *Dark Theme:* Deep slate (`#0E1412`) with vibrant emerald (`#34D399`) and high-contrast text (`#ECEFEC`).
* **Visual Hierarchy:** Clean metric cards, an interactive Altair horizontal bar chart for category breakdown, and responsive `st.dataframe` ledgers.

---

## 🛠️ Integration Details & Architecture
* **Frontend Framework:** `Streamlit` (Layout, state management, native widgets).
* **Data Handling & Calculations:** `Pandas` (Aggregations, group-bys for the Altair charts, formatting ledgers).
* **Visualization:** `Altair` (Declarative statistical visualizations) and `Pillow (PIL)` (Custom geometric vector-style graphics engine replacing traditional desktop `turtle` for cloud compatibility).
* **State Persistence:** `st.session_state` (Ensures that transactions remain strictly isolated to the user's current browser session, preventing multi-device sync collisions in cloud environments).

---

## 🚀 Deployment Instructions

This application is designed for seamless, free deployment on **Streamlit Community Cloud**.

### Local Setup
1. Clone this repository: `git clone https://github.com/YourUsername/YourRepository.git`
2. Navigate to the project directory: `cd YourRepository`
3. Create a virtual environment: `python -m venv venv`
4. Activate the environment:
   * Windows: `venv\Scripts\activate`
   * Mac/Linux: `source venv/bin/activate`
5. Install dependencies: `pip install -r requirements.txt`
6. Run the app: `streamlit run app.py`

### Cloud Deployment (Streamlit Community Cloud)
1. Push your code, including `app.py` and `requirements.txt`, to your GitHub repository.
2. Go to [share.streamlit.io](https://share.streamlit.io/) and log in with your GitHub account.
3. Click **"Create app"** -> **"I already have an app"**.
4. Select your repository, set the branch to `main`, and the Main file path to `app.py`.
5. Click **"Deploy!"**. Streamlit will automatically read the `requirements.txt` and launch the app within 1-2 minutes.

---

## 📸 Application Screenshots

<img width="1920" height="1080" alt="image" src="https://github.com/user-attachments/assets/aadfdb4c-8158-47d6-833c-b1abc8b395f8" />

<img width="1920" height="1078" alt="{8F19112C-3B9F-4573-B133-7D2E4CC9BA34}" src="https://github.com/user-attachments/assets/bcb98c9a-e128-48d8-9baf-3550757827ca" />

<img width="1920" height="1078" alt="{1A6BA418-F0C6-4546-9AE1-8DE7E6030A09}" src="https://github.com/user-attachments/assets/20c18c4c-755c-4696-80e3-f3bc58b36ba0" />

<img width="1920" height="1080" alt="{F162F8DA-93C1-49B6-AC3D-554C6D11C6EB}" src="https://github.com/user-attachments/assets/a2d25c09-13d5-4634-801c-e39af70bf315" />

---
