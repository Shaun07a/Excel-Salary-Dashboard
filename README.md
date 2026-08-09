# Excel Salary Dashboard

![Excel Salary Dashboard](1_Salary_Dashboard.png)

## Introduction

This data jobs salary dashboard was created to help job seekers investigate salaries for their desired jobs and understand how factors such as job title, country, and employment type can influence compensation.

The project was created as part of my Excel Data Analytics learning journey, where I developed practical skills in analyzing, visualizing, and presenting data using Microsoft Excel. The dataset contains detailed information on job titles, salaries, locations, job schedule types, and essential skills.

### Dashboard File

My final dashboard is in [1_Salary_Dashboard.xlsx](1_Salary_Dashboard.xlsx).

### Excel Skills Used

The following Excel skills were utilized for analysis:

- **Charts**
- **Formulas and Functions**
- **Data Validation**

### Data Jobs Dataset

The dataset used for this project contains real-world data science job information from 2023. It includes detailed information on:

- **Job titles**
- **Salaries**
- **Locations**
- **Job schedule types**
- **Skills**

## Dashboard Build

### Charts

#### Data Science Job Salaries - Bar Chart

![Data Science Job Salaries](Salary_Dashboard_Chart1.png)

- **Excel Features:** Utilized the bar chart feature with formatted salary values to present salary comparisons clearly.
- **Design Choice:** Used a horizontal bar chart to make it easier to compare median salaries across different job titles.
- **Data Organization:** Job titles were organized by salary to improve readability and comparison.
- **Insights Gained:** The visualization provides a quick comparison of salary levels across different data-related job roles and highlights higher-paying positions.

#### Country Median Salaries - Map Chart

![Country Median Salaries](Salary_Dashboard_Country_Map.png)

- **Excel Features:** Utilized Excel's Map Chart feature to visualize median salaries across different countries.
- **Design Choice:** Used a geographic visualization to make differences in salary levels between countries easier to identify.
- **Data Representation:** Median salaries were plotted for countries with available salary information.
- **Visual Enhancement:** The map provides an intuitive way to understand geographical salary differences.
- **Insights Gained:** The visualization highlights global salary disparities and provides a quick overview of salary levels across different regions.

### Formulas and Functions

#### Median Salary by Job Titles

```excel
=MEDIAN(
IF(
    (jobs[job_title_short]=A2)*
    (jobs[job_country]=country)*
    (ISNUMBER(SEARCH(type,jobs[job_schedule_type])))*
    (jobs[salary_year_avg]<>0),
    jobs[salary_year_avg]
)
)
