# 🎓 Recent Graduates Employment & Earnings Analysis

An Excel-based education analytics project transforming recent graduate data into actionable insights on employment outcomes, unemployment risk, earnings, gender representation, low-wage employment, and career opportunities across college majors.

This project analyzes **173 college majors** to understand how graduate outcomes differ across fields of study, which majors and major categories experience greater employment challenges, and where earnings disparities or career risks may exist.

The project was completed as my **first milestone project with Data In Motion** and was designed to strengthen a realistic analytics workflow: interpreting business questions, selecting the correct metrics, validating calculations in Excel, creating storytelling visualizations, and translating findings into leadership recommendations.

---

## 📌 Business Problem

Higher education leaders need more than graduate enrollment totals to understand whether students are successfully transitioning into the workforce.

They need to understand which majors experience stronger or weaker employment outcomes, how earnings vary across fields of study, whether larger programs produce stronger career outcomes, and where students may need additional career support.

This project addresses several core areas of graduate outcomes:

### 1. Enrollment & Employment Outcomes

Leadership wants to understand whether popular majors also produce strong employment outcomes.

Final Business Problem:  
Analyze graduate enrollment and employment outcomes to determine whether majors with larger student populations are associated with stronger full-time employment opportunities.

### 2. Unemployment Risk

Leadership wants to identify majors and major categories where graduates may face greater difficulty transitioning into employment.

Final Business Problem:  
Analyze unemployment rates across majors and major categories to identify areas experiencing elevated unemployment and determine where targeted career support may be needed.

### 3. Earnings Outcomes

Leadership wants to understand how graduate earnings differ across fields of study.

Final Business Problem:  
Analyze median earnings and earnings variability across majors to identify fields associated with stronger earnings outcomes and majors where graduate earnings vary significantly.

### 4. Gender Representation & Earnings

Leadership wants to understand whether predominantly male and predominantly female majors experience different earnings outcomes.

Final Business Problem:  
Compare the earnings of majors with more than 70% male or female representation to determine whether meaningful differences exist between the two groups.

### 5. Employment Quality

Leadership wants to understand whether graduates are finding higher-quality employment rather than only measuring whether they are employed.

Final Business Problem:  
Analyze low-wage and part-time employment across majors to identify fields where graduates may face weaker employment quality or earning potential.

---

## 🎯 Project Objective

The objective of this analysis was to transform recent graduate data into a structured decision-making tool that could help higher education leadership better understand employment and earnings outcomes across college majors.

The analysis focused on questions such as:

- Which majors have the largest graduate populations?
- Which major categories experience the highest unemployment rates?
- Do larger majors produce stronger full-time employment outcomes?
- How do earnings differ between predominantly male and predominantly female majors?
- Which majors experience the greatest earnings variability?
- Which areas may require additional career-development support?

The project also aimed to move beyond identifying only the highest or lowest values by explaining what the findings mean and connecting them to measurable leadership recommendations.

### Business Questions Answered

1. What are the 5 most common majors among recent graduates?
2. What is the gender distribution across these 5 majors?
3. Which majors have the highest and lowest median earnings?
4. How does the unemployment rate vary across different major categories?
5. Which majors are predominantly female (>70% women), and how do their median earnings compare with predominantly male majors?
6. Are students in more popular majors more likely to find full-time, year-round employment?
7. Is there a relationship between the number of full-time employees and median earnings?
8. For each major category, what is the average number of low-wage jobs?
9. How many majors have an unemployment rate above the hypothetical national average of 6%, and which categories do they belong to?
10. Among the top 10 majors by total enrollment, which has the highest percentage of part-time workers?
11. Which majors have the largest difference between 25th and 75th percentile earnings?
12. What recommendations should leadership consider to improve graduate employment and earnings?
13. Do majors with more employed graduates tend to have higher or lower median earnings?

---

## 📊 Dataset Overview

| Attribute | Description |
|---|---|
| **Rows Analyzed** | 173 majors |
| **Primary Analysis Level** | Major and Major Category |
| **Primary Analysis Areas** | Graduate enrollment, employment, unemployment, earnings, gender representation, low-wage employment |
| **Data Type** | Major-level recent graduate data |
| **Primary Tool** | Microsoft Excel |

### Key Variables

The dataset includes:

- Major
- Major Category
- Total Graduates
- Men
- Women
- ShareWomen
- Employed
- Full-Time
- Full-Time Year-Round
- Part-Time
- Unemployed
- Unemployment Rate
- Median Earnings
- 25th Percentile Earnings
- 75th Percentile Earnings
- Low-Wage Jobs

### Data Quality Notes

The dataset required careful review of the difference between counts, percentages, rates, averages, and categorical fields.

The analysis and validation process included:

- Reviewing missing values before excluding records
- Formatting percentages consistently
- Creating calculated gender-share fields
- Creating employment-rate calculations
- Recalculating category-level unemployment rates
- Avoiding incorrect aggregation of percentage fields
- Creating earnings-spread calculations
- Validating PivotTable results against underlying data
- Confirming that filters and thresholds matched the wording of each business question

A key validation lesson involved unemployment rates. Percentage values could not simply be summed across majors. Category-level unemployment needed to be calculated using total unemployed graduates divided by the corresponding total labor force.

### Limitations

- The project uses a recent-graduate dataset and does not represent every possible factor affecting employment or earnings.
- The analysis identifies relationships and patterns but does not establish causation.
- Earnings are analyzed at the major level rather than individual graduate level.
- The hypothetical 6% unemployment benchmark was provided for analytical comparison and should not be interpreted as an official current national benchmark.
- Additional variables such as geography, industry, occupation, experience, and advanced education could provide further context.

---

## 🛠️ Tools & Technologies

| Tool | Purpose |
|---|---|
| **Microsoft Excel** | Data analysis, formatting, validation, formulas, filtering, sorting, and visualization |
| **Excel Formulas** | Percentage calculations, classification logic, employment rates, averages, and earnings spreads |
| **PivotTables** | Major-category aggregation, benchmark analysis, and summary calculations |
| **Excel Charts** | Bar charts, stacked charts, scatterplots, relationship analysis, and storytelling visualizations |

---

## 🔄 Methodology

Recent Graduate Dataset  
        ↓  
Understand Business Questions  
        ↓  
Select Relevant Variables  
        ↓  
Create Calculated Fields  
        ↓  
Validate Calculations  
        ↓  
Aggregate / Filter Data  
        ↓  
Build Visualizations  
        ↓  
Interpret Findings  
        ↓  
Develop Leadership Recommendations  
        ↓  
Create Executive Summary  

---

## Ask — Define the Business Questions

The analysis began by breaking each business question into the specific analytical components required to answer it.

Clarification focused on:

* Population being analyzed
* Dimensions
* Metrics
* Thresholds
* Comparison groups
* Correct aggregation methods
* Relationships being investigated
* Appropriate visualization types

This helped ensure that each analysis directly answered the business question rather than simply displaying available data.

---

## Prepare — Understand the Data

The dataset contained 173 college majors with employment, earnings, demographic, and graduate-outcome information.

Excel was used to inspect the structure of the data and determine which fields were appropriate for each business question.

Special attention was given to:

* Major-level vs. major-category-level analysis
* Count fields vs. percentage fields
* Employment measures
* Earnings measures
* Gender-share variables
* Missing values
* Percentile-based earnings fields

---

## Process — Calculate & Validate

Several calculated fields were created to support the analysis.

### Share of Men

```excel
=1-ShareWomen
```

### Gender Classification

```excel
=IF(ShareWomen>70%,"Female-Dominated",IF(ShareMen>70%,"Male-Dominated","Neither"))
```

### Full-Time Year-Round Employment Rate

```excel
=Full_Time_Year_Round/Total
```

### Part-Time Percentage

```excel
=Part_Time/Total
```

### Earnings Spread

```excel
=P75th-P25th
```

### Category Unemployment Rate

```excel
=Total_Unemployed/Total_Labor_Force
```

The calculations were validated throughout the project to ensure that percentages, averages, and rates were interpreted correctly.

One important correction involved unemployment rates. Instead of summing unemployment percentages across majors, the category-level rate was recalculated using the total unemployed population divided by the total labor force.

---

## Analyze — Employment & Earnings Analysis

The dataset was analyzed in Excel using formulas, filtering, PivotTables, ranking, and relationship analysis.

The analysis included:

* Ranking majors by total graduate enrollment
* Comparing gender distribution across high-enrollment majors
* Identifying the highest and lowest median earnings
* Calculating unemployment rates by major category
* Comparing predominantly male and female majors
* Investigating enrollment vs. full-time employment
* Investigating full-time employment vs. median earnings
* Calculating average low-wage employment by major category
* Comparing majors against a 6% unemployment benchmark
* Calculating part-time employment percentages
* Measuring earnings variability using 25th and 75th percentile earnings
* Developing leadership recommendations from analytical findings
* Creating an independent relationship question for additional analysis

---

## Share — Data Storytelling

The findings were translated into storytelling visualizations designed to communicate the main takeaway directly through the chart title and visual emphasis.

The workbook includes:

Top 5 Graduate Enrollment Analysis  
Gender Distribution Analysis  
Highest vs. Lowest Median Earnings  
Unemployment by Major Category  
Predominantly Male vs. Female Major Earnings  
Enrollment vs. Full-Time Employment  
Full-Time Employment vs. Median Earnings  
Low-Wage Employment by Major Category  
6% Unemployment Benchmark Analysis  
Part-Time Employment Analysis  
Earnings Range Analysis  
Leadership Recommendations  
Independent Employment vs. Earnings Analysis  

---

## Act — Leadership Recommendations

The final stage converted the analytical findings into specific actions that leadership could evaluate when developing career-support programs and employment initiatives.

Each recommendation was connected to a measurable finding and included a metric that could be monitored after implementation.

---

## 📈 Analysis Preview

### Recent Graduates Employment & Earnings Analysis

The final project summary consolidates the major findings from the 13 business questions into a single visual analysis.

The complete workbook contains the calculations, supporting data, PivotTables, charts, and leadership recommendations used throughout the project.

---

## 🔍 Key Findings

### 1. Psychology combines the largest graduate population with relatively high unemployment

**What happened:**  
Psychology was the most common major with approximately **393.7K graduates**.

Among the top 5 most common majors, Psychology also had the **highest unemployment rate at 8.38%**.

**Why it matters:**  
High enrollment does not automatically translate into stronger employment outcomes.

**Business implication:**  
Leadership should evaluate employment outcomes for high-enrollment majors rather than assuming popularity reflects stronger career opportunities.

### 2. Social Science has the highest unemployment rate by major category

**What happened:**  
Social Science had an unemployment rate of **9.67%**, compared with **4.95% for Education**.

**Why it matters:**  
Social Science graduates appear to experience greater difficulty transitioning into employment than graduates in several other major categories.

**Business implication:**  
Career-support resources may need to be more targeted toward categories experiencing elevated unemployment.

### 3. Above-benchmark unemployment is widespread

**What happened:**  
Approximately **68% of majors** exceeded the hypothetical **6% unemployment benchmark**.

Engineering accounted for **29 majors** above the benchmark.

**Why it matters:**  
Above-benchmark unemployment is not isolated to only a small number of majors.

**Business implication:**  
Leadership can use the benchmark to identify higher-risk majors and determine where additional career support may be most useful.

### 4. Predominantly male majors have higher average median earnings

**What happened:**  
Predominantly male majors averaged approximately **$52,094** in major-level median earnings.

Predominantly female majors averaged approximately **$33,162**.

This represents an earnings difference of approximately **$18.9K**.

**Why it matters:**  
The earnings outcomes of heavily male- and female-represented majors differ substantially in this dataset.

**Business implication:**  
Leadership should investigate opportunities to strengthen pathways into higher-paying careers for students in predominantly female majors.

### 5. Communications & Journalism has the highest average number of low-wage jobs

**What happened:**  
Communications & Journalism had an average of approximately **12.4K low-wage jobs**.

Agriculture & Natural Resources had approximately **790**.

**Why it matters:**  
Employment alone does not fully describe graduate outcomes. Job quality and earning potential also matter.

**Business implication:**  
Programs with high low-wage employment may benefit from stronger alignment with in-demand skills and higher-paying career pathways.

### 6. Enrollment is not strongly associated with higher full-time employment rates

**What happened:**  
The relationship between total enrollment and full-time year-round employment rates was weak.

**Why it matters:**  
Larger majors do not necessarily produce stronger full-time employment outcomes.

**Business implication:**  
Employment success should be evaluated separately from program popularity.

### 7. More full-time employees do not necessarily mean higher earnings

**What happened:**  
Majors with more full-time employees showed a slight negative relationship with median earnings.

**Why it matters:**  
Employment volume and earning potential measure different aspects of graduate success.

**Business implication:**  
Leadership should evaluate both job placement and earnings instead of relying on employment counts alone.

### 8. Astronomy & Astrophysics has the widest earnings range

**What happened:**  
Astronomy & Astrophysics had a **$77.5K gap between the 25th and 75th percentile earnings**, the largest among the majors analyzed.

Pharmaceutical Sciences followed with a **$70K gap**.

**Why it matters:**  
Graduates within the same major can experience significantly different earnings outcomes.

**Business implication:**  
Large earnings ranges warrant deeper investigation into the career paths, industries, or specializations associated with stronger outcomes.

---

## 💼 Business Impact

This project demonstrates how higher-education data can be converted into a structured framework for evaluating graduate career outcomes.

The analysis can help leadership:

* Identify majors and major categories experiencing higher unemployment.
* Evaluate employment outcomes independently from enrollment popularity.
* Identify earnings differences across groups of majors.
* Monitor low-wage and part-time employment outcomes.
* Identify majors with significant earnings variability.
* Target career-development programs toward areas showing greater employment risk.
* Measure whether career-support initiatives improve graduate outcomes over time.

📌 In short: this project turns recent graduate data into a decision-making framework for identifying employment challenges, understanding earnings differences, and improving career pathways for students.

---

## 🚀 Recommendations

Recommendation | Why It Matters
---|---
Strengthen employment support for Social Science students | Social Science has the highest unemployment rate at 9.67%
Expand pathways into higher-paying careers for predominantly female majors | Predominantly male majors average approximately $18.9K more in median earnings
Strengthen in-demand skill development for Communications & Journalism | Communications & Journalism has the highest average number of low-wage jobs
Target career support toward majors exceeding the 6% unemployment benchmark | Approximately 68% of majors exceed the benchmark
Investigate career paths driving large earnings differences | Astronomy & Astrophysics has a $77.5K gap between the 25th and 75th percentile earnings

---

## 🔮 Next Steps

This analysis provides a foundation for continued graduate employment and earnings analysis.

Potential extensions include:

* Industry-level analysis — determine which industries graduates from each major enter.
* Occupation analysis — identify job roles associated with stronger employment and earnings outcomes.
* Geographic analysis — evaluate whether employment and earnings outcomes differ across regions.
* Career-path analysis — investigate which career paths contribute to large earnings differences within majors.
* Gender earnings analysis — investigate individual-level gender earnings differences if appropriate data becomes available.
* Benchmark monitoring — compare graduate unemployment rates against updated labor-market benchmarks.
* Interactive dashboard development — recreate the analysis in Tableau or Power BI for deeper stakeholder exploration.
* Longitudinal analysis — evaluate how graduate outcomes change across multiple graduating classes or time periods.

---

## 📁 Project Structure

```text
Recent-Graduates-Employment-Earnings-Analysis/
│
├── Recent_Graduates_Questions_1_13.png
├── Recent_Graduates_Employment_Earnings_Analysis.xlsx
└── README.md
```

---

## ⭐ Skills Demonstrated

Business Analysis  
Microsoft Excel  
Data Analysis  
Data Validation  
PivotTables  
Excel Formulas  
Calculated Fields  
Data Visualization  
Data Storytelling  
Employment Analysis  
Earnings Analysis  
Benchmark Analysis  
Percentage & Rate Calculations  
Scatterplot Analysis  
Stakeholder Communication  
Leadership Recommendations  
Critical Thinking  
Problem Solving  

---

## 📚 Key Takeaways

This project reinforced that strong analysis begins with understanding the business question before selecting metrics or building a visualization.

One of the biggest lessons was understanding the difference between counts, percentages, averages, and rates. For example, unemployment percentages could not simply be summed across majors. Category unemployment needed to be recalculated using total unemployed graduates divided by the corresponding labor force.

The project also strengthened my ability to move beyond simply identifying the highest or lowest value. A useful finding should explain what happened, what the result means, and how a stakeholder could respond.

Another important takeaway was the value of validating the scope of an analysis. When a business question specifies a threshold such as more than 70% female or male representation, every qualifying major should be evaluated rather than relying on only a smaller sample.

Finally, the project reinforced the importance of data storytelling. A visualization should not simply display data. The chart type, title, comparison, formatting, and supporting explanation should work together to communicate the takeaway clearly.

The project demonstrates the analytical workflow:

**Business Question → Analysis → Validation → Visualization → Insight → Recommendation**
