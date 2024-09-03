# Excel-Case-Study_01
### **Project Report: Net Revenue Management Case Study for HealthMax**

---

#### **Introduction**

As a Category Manager for HealthMax, a leading company in the shampoo industry, the goal was to perform a detailed market analysis and explore Net Revenue Management (NRM) opportunities. This project utilized advanced Excel techniques to clean data, analyze market trends, and provide strategic recommendations. The following report outlines the steps taken and the key findings from this analysis.

---

#### **Data Cleaning and Preparation**

1. **Value Formatting:**
   - Values were separated into distinct columns for better organization.
   - Removed hardcoded dollar signs from the "Values Month" column and reformatted the column into currency.

2. **PivotTable Creation:**
   - Created a PivotTable to overview different brands per supplier and renamed the worksheet to "Brands per Supplier."
   - Another PivotTable was created in a new worksheet, "HealthMax Growth," displaying value sales per brand for all years, filtering on HealthMax as the supplier. Growth numbers compared to the previous year were calculated, excluding 2023 data due to its limited timeframe.

3. **YTD and MAT Calculations:**
   - Added a "Units YTD" column and used the SUMIFS() function to calculate unit sales by brand, region, and year.
   - Calculated "Values YTD" using the SUMIFS() formula and applied currency formatting.
   - Added "Units MAT" and "Values MAT" columns, calculating them with additional SUMIFS() functions to account for the corresponding values from the previous year.

4. **Turnover Calculation:**
   - A PivotTable was created to calculate the total turnover (revenue) of the shampoo category using MAT values, focusing on March 2023. The worksheet was renamed "MAT Value Total Category."

---

#### **Market Analysis**

1. **Market Share Analysis:**
   - Created a PivotTable to calculate total value sales per year, per brand, which was compiled in a new worksheet named "Market Share."
   - This analysis allowed for tracking brand performance over time, highlighting key trends and shifts in market share.

2. **Gross Profit Analysis:**
   - Added columns "Gross Profit per unit" and "Gross Profit per product," calculating HealthMax's profitability across its product portfolio. Currency formatting was applied for clarity.
   - Summarized gross profit data to ensure accurate financial reporting.

3. **Net Sales Contribution:**
   - A "Net Sales Contribution" column was added, calculating the contribution of each product to overall sales in 2022. The results were formatted as percentages to provide clear insights.

4. **Profitability Matrix:**
   - Created a PivotTable in a new worksheet, "Profitability Matrix," organizing data by Gross Margin and Net Sales Contribution. This matrix provided a visual representation of product profitability, assisting in identifying high-performing products.

---

#### **Net Revenue Management (NRM) Opportunities**

1. **NRM Pillars:**
   - Gained a comprehensive understanding of the five pillars of NRM within the FMCG industry.
   - Successfully translated these pillars into actionable business opportunities for HealthMax, focusing on optimizing revenue and profitability.

2. **Growth Identification:**
   - A PivotTable was created in the "New Category Opportunity" worksheet to identify the fastest-growing subcategory by comparing unit growth from 2018 to 2022.
   - This analysis revealed potential areas for expansion and strategic investment.

---

#### **Recommendations and Strategic Implementation**

1. **Forecasting and Strategy:**
   - Forecasted net sales for the coming year, demonstrating the potential impact of NRM initiatives on growth.
   - Developed strategic recommendations for management, outlining clear steps to increase net sales and profitability.

2. **Implementation Roadmap:**
   - Prepared a detailed implementation plan, focusing on leveraging NRM opportunities to unlock extra potential in the market.

---

#### **Conclusion**

This project provided valuable insights into the shampoo category within the FMCG industry. By leveraging Excel's advanced functions, including PivotTables, SUMIFS(), and financial metrics calculations, the analysis identified key growth opportunities and strategic initiatives. The findings and recommendations from this study can guide HealthMax in optimizing its market presence and achieving sustained revenue growth.

---
