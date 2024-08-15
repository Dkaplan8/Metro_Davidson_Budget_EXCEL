Metro Budget Analysis Project

Overview:

This project focuses on analyzing the budget and actual spending of various departments over multiple years. The objectives include calculating differences and percentage differences, ranking departments based on their budget performance, and utilizing Excel functions such as IFERROR, VLOOKUP, XLOOKUP, INDEX, MATCH, and data validation. The project aims to provide a comprehensive understanding of the financial performance of departments, enabling better budget management and decision-making.

Learning Objectives:

Write formulas using IFERROR or IF to handle errors.
Correctly utilize absolute cell references to enable formula copying.
Perform two-way lookups using VLOOKUP, XLOOKUP, and INDEX + MATCH.
Utilize data validation to create dropdown menus.

Project Steps:

1. Calculating Differences and Percentage Differences
Difference Calculation: Compute the difference between the actual spending and the budget for each department using the formula =Actual - Budget.
Percentage Difference: Calculate the percent difference by dividing the difference by the budget using the formula =(Actual - Budget) / Budget.
Error Handling: Use IFERROR to handle potential division by zero errors in the percentage difference calculation.

2. Ranking Departments
Rank Calculation: Rank each department based on their percentage difference for each year. Departments with the highest percentage of their budget remaining should be ranked number 1.
Error Handling: Address any #DIV/0! errors in the rank column by using IFERROR or a logical function.

3. Retrieving Values Using VLOOKUP
VLOOKUP Application: Fill in a table to retrieve the difference for selected departments for each year using VLOOKUP.

4. Retrieving Values Using XLOOKUP
XLOOKUP Application: Repeat the retrieval process using XLOOKUP, ensuring the formula can be copied down but not necessarily across.

5. Retrieving Values Using INDEX and MATCH
INDEX + MATCH Application: Use INDEX and MATCH to retrieve values for the selected departments.

6. Data Validation and Dynamic Table
Data Validation: Create a dropdown menu in cell B82 to select a department using data validation.
Dynamic Retrieval: Use INDEX and MATCH to fill a table with the budget and actual spending for each financial year based on the selected department.


Conclusion:

This project provides a structured approach to analyzing and understanding the budget performance of departments using various Excel functions and tools. By calculating differences, ranking performance, and utilizing advanced lookup functions, we can gain valuable insights into financial management and decision-making. The use of data validation and dynamic tables further enhances the interactivity and usability of the analysis, making it a comprehensive tool for budget analysis.

For any questions or further assistance, please contact me:

Project Maintainer: Douglas Kaplan
Contact Information: Douglasjkaplan@gmail.com
Date: 5/21/24
