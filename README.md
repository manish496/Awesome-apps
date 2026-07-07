# App Performance Data Analysis (Excel)

This repository contains an Excel-driven data analysis project focused on mobile app performance tracking, user engagement, and download patterns. The analysis answers core business questions regarding application growth, retention metrics, and rating distributions using advanced Excel functions.

## 📊 Project Overview

The project evaluates monthly data across several mobile applications (e.g., *Sugar Rush*, *Chocolate Crossing*, *Subway Wafers*, *Crushed by Candies*) over a specified timeline in 2021. It serves as a practical showcase for transforming raw app metrics into actionable business intelligence insights.

### Key Business Questions Addressed
* **Volume Tracking:** Calculating cumulative download metrics for specific apps and app categories (e.g., apps starting with "C").
* **User Churn & Retention:** Measuring high-churn periods by isolating uninstalls within specific months and identifying maximum monthly uninstall ratios.
* **Milestone Benchmarking:** Tracking how frequently apps surpassed high-performance benchmarks (e.g., >8,000 downloads/month).
* **Sentiment Analysis:** Analyzing user feedback ratios, identifying apps with the highest 1-star distributions, and monitoring shifts in top-tier (5-star) review rates over time.
* **Extremum Performance:** Pinpointing ranked historical highs (1st, 2nd, and 3rd place achievements) across months and products.

---

## 🛠️ Excel Toolset & Formulas Used

Instead of relying on basic lookups, this repository showcases advanced logical, statistical, and array formulas optimized for deep-dive analysis:

*   **Conditional Sums & Counts:** `SUMIFS`, `COUNTIFS` *(Multi-criteria transactional filtering)*
*   **Averages & Outliers:** `AVERAGEIFS`, `MAXIFS`, `MINIFS` *(Targeted performance boundaries)*
*   **Statistical Ranking:** `LARGE`, `SMALL` *(Isolating top tier rankings and lowest outliers)*
*   **Dynamic Lookups:** `XLOOKUP`, `INDEX / MATCH` *(Flexible bi-directional array cross-referencing)*
*   **Data Reshaping:** `FILTER` *(Dynamic sub-dataset extraction)*

---

## 📁 Repository Structure

*   `top-formulas-for-data-analysis-answers.xlsx`: The core workbook featuring:
    *   `Questions`: A structured breakdown of the business analytical requirements.
    *   `AA`: The main data repository housing raw metrics (Downloads, Uninstalls, 5-Star Ratings, 1-Star Ratings, Total Ratings) paired with the live functional formulas resolving the business queries.
