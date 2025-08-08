# ExhausOLS

ExhausOLS is a **user-friendly Python-based GUI** for performing **exhaustive OLS regression**. It helps identify the **optimal combination of predictor variables** by testing all possible model combinations. The tool is designed for researchers and analysts who want to explore variable selection in Ordinary Least Squares (OLS) regression through a fully automated process.

---

## 🔽 How to Use

**No installation required – just open the `ExhausOLS.exe` located in the `dist` folder and follow the steps.**

1. **Download and open the software.**  
   *(Note: On the first run, it may take a few seconds to launch.)*

2. **Click the "Browse" button** and upload your dataset.  
   The input file must be an **Excel (.xlsx)** file that includes:
   - One column for the **dependent variable**
   - One or more columns for the **independent variables**

   Example:

   | Y   | X1  | X2  | X3  |
   |-----|-----|-----|-----|
   | 4.2 | 1.3 | 5.1 | 2.8 |
   | 3.9 | 0.7 | 4.8 | 3.1 |

   *(All columns must contain **numeric values**.)*

3. **Select your dependent variable** from the dropdown list.

4. **Select one or more independent variables** to include in the analysis.

5. **Click the "Run" button** to start the analysis.

6. When prompted, **choose a location and file name** to save the results.

7. **Wait for the process to complete.**  
   *(The runtime depends on the size of your dataset and the number of selected variables.)*

8. Once finished, a **CSV file** will be created containing:
   - All possible combinations of the selected independent variables  
   - Regression results for each model, including the following columns:

     - `Model_ID`: Unique identifier for each model  
     - `Num_Variables`: Total number of predictors in the model  
     - `Num_Significant`: Number of statistically significant predictors (p < 0.05)  
     - `Variables`: List of all predictors used  
     - `Significant_Variables`: Subset of predictors with p-values < 0.05  
     - `Adjusted_R2`: Adjusted R-squared of the model  
     - `AIC`: Akaike Information Criterion  
     - `AICc`: AIC corrected for small samples  
     - `BIC`: Bayesian Information Criterion  
     - `VIFs`: List of Variance Inflation Factor values for predictors  
     - `P_Values`: List of p-values for predictors  

   You can use this file to identify the most suitable variable combination for your analysis.

---

## 📁 Contents

This repository includes the complete source code, compiled executable, and sample dataset for full transparency and reproducibility.

---

## 🔗 Citation

If you use this tool in research, please cite the following:

Zarei, M., Rafieian, M., and Soltani, A. Exploring spatial disparities in bike-sharing usage: Lessons from the case of Tehran.  
Journal xxxx  
DOI: xxxx
GitHub repo: https://github.com/M-Zarei/ExhausOLS

---

## 📄 License

Released under the [MIT License](LICENSE).
