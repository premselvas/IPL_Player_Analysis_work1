# IPL_Player_Analysis_work1
# 🏏 IPL Player Data Analysis using Excel

A comprehensive IPL (Indian Premier League) player performance analysis project built using Microsoft Excel. This project analyzes batting and bowling statistics of 559 IPL players from different countries to identify top performers, aggressive batsmen, economical bowlers, all-rounders, and country-wise cricket insights.

The project demonstrates practical data analysis techniques using Excel formulas, filtering, sorting, pivot tables, and statistical analysis.

---

# 📌 Project Overview

The Indian Premier League (IPL) is one of the world’s most competitive T20 cricket tournaments. This project focuses on analyzing IPL player statistics to extract meaningful insights regarding batting, bowling, and overall player performance.

The analysis includes:

- 📊 Country-wise player statistics
- 🏏 Batting performance analysis
- 🎯 Bowling performance analysis
- 🔥 Aggressive batsmen identification
- 🧤 Best wicket-taking bowlers
- ⚡ Economy rate analysis
- 🧠 All-rounder performance detection
- 📈 Excel-based data analytics and visualization

---

# 🚀 Features

| Feature | Description |
|---|---|
| 📊 Country-wise Analysis | Compare player statistics across countries |
| 🏏 Batting Analysis | Analyze runs, averages, strike rates |
| 🎯 Bowling Analysis | Evaluate wickets and economy rates |
| 🔥 Aggressive Batsmen Detection | Identify high strike-rate players |
| 🧤 Best Bowlers Identification | Find top wicket-takers |
| 🧠 All-Rounder Detection | Players contributing in batting & bowling |
| 📈 Excel Formulas & Functions | SUM, SUMIF, COUNTIF, AVERAGE, Filters |
| 📋 Pivot Table Support | Dynamic data exploration |
| 📉 Statistical Insights | Performance comparisons and trends |

---

# 📊 Dataset Summary

The dataset contains statistics of **559 IPL players** from multiple cricket-playing nations.

| Country | Player Count |
|---|---|
| India | 257 |
| Australia | 72 |
| South Africa | 39 |
| New Zealand | 22 |
| Sri Lanka | 20 |
| West Indies | 19 |
| England | 14 |
| Pakistan | 13 |
| Bangladesh | 5 |
| Zimbabwe | 2 |
| Netherlands | 1 |
| **Total Players** | **559** |

---

# 📈 Overall Tournament Statistics

## 🏏 Batting Statistics

| Metric | Total |
|---|---|
| Total Balls Played | 173,193 |
| Total Runs Scored | 222,203 |
| Total Strike Rate (Sum) | 54,403.66 |
| Total Batting Average (Sum) | 7,883.15 |

---

## 🎯 Bowling Statistics

| Metric | Total |
|---|---|
| Total Balls Delivered | 179,078 |
| Total Runs Conceded | 222,203 |

---

# 📋 Analysis Performed

| Sl.No | Analysis | Method Used |
|---|---|---|
| 1 | Total Runs Scored in Tournament | SUM |
| 2 | Total Wickets Taken | SUM |
| 3 | Total Runs by Team | SUMIF |
| 4 | Average Strike Rate | AVERAGE |
| 5 | Average Runs Per Player | Runs / Matches |
| 6 | Top Wicket-Taker | MAX |
| 7 | Lowest Economy Rate | MIN |
| 8 | Players with SR > 100 | COUNTIF |
| 9 | Top Scoring Players | Sorting |
| 10 | Best Bowlers | Sorting |
| 11 | Aggressive Batsmen | Strike Rate Analysis |
| 12 | Economical Bowlers | Economy Rate Analysis |
| 13 | High Batting Performance Players | Filters |
| 14 | Top Wicket-Takers | Filters |
| 15 | India & Australia Player Analysis | Filters |
| 16 | All-Rounders Identification | Multiple Conditions |
| 17 | Aggressive Batsmen Display | SR > 120 |
| 18 | Economical Bowlers Display | Economy < 5 |

---

# 📁 Dataset Columns

| Column Name | Description |
|---|---|
| Player ID | Unique player identifier |
| Player Name | Cricketer name |
| DOB | Date of birth |
| Country | Represented nation |
| Balls Played | Balls faced by batsman |
| Runs Scored | Total batting runs |
| Batting Average | Average runs per dismissal |
| Strike Rate | (Runs / Balls) × 100 |
| Balls Delivered | Balls bowled |
| Runs Conceded | Runs given while bowling |
| Economy Rate | Runs conceded per over |

---

# 🔍 Key Insights

## 🏏 Top Batting Nations

- 🇮🇳 India dominates with **257 players** and **123,842 runs**
- 🇦🇺 Australia ranks second with **32,437 runs**
- 🇿🇦 South Africa stands third with **21,079 runs**

---

## 🎯 Top Bowling Nations

- 🇮🇳 India delivered **104,798 balls**
- 🇦🇺 Australia delivered **18,578 balls**
- 🇿🇦 South Africa delivered **13,627 balls**

---

# 🌟 Legendary Players Included

The dataset contains performance statistics of famous IPL players such as:

- **Virat Kohli** 🇮🇳 — 5,426 runs
- **Rohit Sharma** 🇮🇳 — 4,902 runs
- **MS Dhoni** 🇮🇳 — 4,450 runs
- **David Warner** 🇦🇺 — 4,717 runs
- **Chris Gayle** 🇯🇲 — 4,525 runs
- **AB de Villiers** 🇿🇦 — 4,414 runs

---

# 📂 Excel Sheet Structure

| Sheet Name | Description |
|---|---|
| Sheet1 | Country-wise player count |
| Sheet2 | Country-wise balls delivered |
| Sheet3 | Country-wise batting analysis |
| Sheet4 | Strike rate & batting average |
| Dataset | Complete IPL player dataset |
| Instructions | Task instructions |
| Solution Sheet | Formula solutions & methods |

---

# 🛠️ Excel Functions Used

| Function | Purpose |
|---|---|
| SUM | Total calculations |
| SUMIF | Conditional summation |
| COUNTIF | Conditional counting |
| AVERAGE | Mean calculations |
| SORT | Ranking players |
| FILTER | Extracting player groups |
| Pivot Tables | Dynamic analysis |

---

# 🖥️ How to Use

## 🔹 Total Runs by India

```excel
=SUMIF(C2:C560,"India",E2:E560)
```

---

## 🔹 Players with Strike Rate > 100

```excel
=COUNTIF(G2:G560,">100")
```

---

## 🔹 Economy Rate Formula

```excel
=Runs_Conceded/(Balls_Delivered/6)
```

---

## 🔹 Average Strike Rate

```excel
=AVERAGE(StrikeRate_Column)
```

---

# 🔎 Filtering Examples

## 🧠 Find All-Rounders

```text
Runs Scored > 0 AND Wickets Taken > 0
```

---

## 🔥 Find Aggressive Batsmen

```text
Strike Rate > 120
```

---

## 🎯 Find Economical Bowlers

```text
Economy Rate < 5.0
```

---

## 🇮🇳 Filter Indian Players

```text
Country = India
```

---

# 📊 Sample Analysis Queries

## 🏏 Top 5 Run Scorers

Sort the dataset by:
- Runs Scored → Descending Order

---

## 🎯 Most Economical Bowlers

Filter:
- Balls Delivered > 300

Then sort:
- Economy Rate → Ascending Order

---

## 🧠 Best All-Rounders

Filter:
- Runs Scored > 500
- Wickets Taken > 15

---

## 📈 Country Performance Comparison

Use Pivot Tables to compare:
- Average Strike Rate
- Average Economy
- Runs Per Player
- Wickets Per Player

---

# 💡 Advanced Insights

This project can also be extended to identify:

- 🏆 Most Valuable Players (MVPs)
- ⚡ Power Hitters
- 🎯 Death Over Specialists
- 🧤 Bowling Dominators
- 🌍 Foreign Player Impact
- 📊 Team Balance Analysis
- 📉 Age vs Performance Trends

---

# 📌 Important Notes

- Some rows may contain blank values
- Strike Rate Formula:
  
```text
(Runs Scored / Balls Played) × 100
```

- Economy Rate Formula:

```text
Runs Conceded / (Balls Delivered / 6)
```

- Batting Average Formula:

```text
Runs Scored / Number of Dismissals
```

---

# 🔧 Future Enhancements

- ✅ Year-wise IPL Analysis
- ✅ Dashboard Creation using Excel Charts
- ✅ Player Comparison System
- ✅ Venue-based Performance Analysis
- ✅ Bowling Average & Strike Rate
- ✅ Interactive Power BI Dashboard
- ✅ Advanced Statistical Analysis
- ✅ Machine Learning-based Predictions

---

# 🛠️ Technologies Used

| Technology | Purpose |
|---|---|
| Microsoft Excel | Data Analysis |
| Pivot Tables | Data Summarization |
| Excel Functions | Statistical Calculations |
| Charts & Graphs | Data Visualization |

---

# 👨‍💻 Developer

**Prem Selva S**  
B.Tech Artificial Intelligence and Data Science  
2026 Batch  

Project: **IPL Player Data Analysis using Excel**

---

# ⭐ GitHub Support

If you found this project useful:

⭐ Star the repository  
🍴 Fork the project  
📊 Share your feedback

---

# 📜 License

This project is created for educational and analytical purposes only.
