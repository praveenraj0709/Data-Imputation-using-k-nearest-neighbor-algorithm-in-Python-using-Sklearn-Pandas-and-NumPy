# Data Imputation using k-nearest neighbor algorithm in Python using Sklearn, Pandas and NumPy
In this project, we perform missing data imputation in Python using 2 variants of the KNN algorithm, i.e Complete case KNN and Incomplete case KNN, using Scikit Learn, Pandas and NumPy. 

Here, we compare and analyze the efficiency of these two approaches by imputing missing values in [**26 datasets**](https://github.com/Dalersingh-rs/Data-Imputation-using-k-nearest-neighbor-algorithm-in-Python-using-Sklearn-Pandas-and-NumPy/blob/main/List%20of%20Datasets.xlsx) **(numerical-16, categorical-5, mixed-5)**, having **44** combinations of missing values each (**1144** datasets in total). The number of instances in a dataset range from **150** to **4,94,020**.

### Libraries Required

1. Pandas
2. Math
3. Os
4. Glob
5. Numpy
6. Sklearn
7. Openpyxl
8. Categorical_Encoders

### Code Functionality

Pandas is used to read and identify the missing values in the dataset. 

KNN Imputer module in Scikit Learn is used and adjusted according to the functioning of both the algorithms to impute the missing values in the data.

NumPy is used to calculate the NRMS and AE values by comparing the original and imputed datasets, and Pandas is used to write these values in an [excel file](https://github.com/Dalersingh-rs/Data-Imputation-using-k-nearest-neighbor-algorithm-in-Python-using-Sklearn-Pandas-and-NumPy/blob/main/Empty%20Table-NRMS-AE.xlsx) for future comparisons

Efficiency is analyzed by comparing and visualizing the NRMS/AE values of individual datasets in both algorithms.

>To get more insight on how the algorithms function, checkout the [research paper](https://github.com/Dalersingh-rs/Data-Imputation-using-k-nearest-neighbor-algorithm-in-Python-using-Sklearn-Pandas-and-NumPy/blob/main/2014%20Jason%20Van%20Hulse%20-%20Incomplete%20case%20nearest%20neighbor%20imputation%20in%20sof%20%5Bretrieved_2022-01-28%5D.pdf) that I used

### File Description 
To simplify execution with a better runtime, 5 python files have been attached. 

To impute numerical datasets, run

1. [Numerical-incomplete_case.py](https://github.com/Dalersingh-rs/Data-Imputation-using-k-nearest-neighbor-algorithm-in-Python-using-Sklearn-Pandas-and-NumPy/blob/main/Numerical/Numerical-%20Incomplete_Case.py)(preferred because of Efficient Runtime) 
2. [Numerical-complete_case.py](https://github.com/Dalersingh-rs/Data-Imputation-using-k-nearest-neighbor-algorithm-in-Python-using-Sklearn-Pandas-and-NumPy/blob/main/Numerical/Numerical%20-%20Complete_Case.py)

To impute Categorical datasets,run

3. [Categorical-incomplete_case.py](https://github.com/Dalersingh-rs/Data-Imputation-using-k-nearest-neighbor-algorithm-in-Python-using-Sklearn-Pandas-and-NumPy/blob/main/Categorical/Categorical-%20Incomplete_case.py)(preferred because of Efficient Runtime) 
4. [Categorical-complete_case.py](https://github.com/Dalersingh-rs/Data-Imputation-using-k-nearest-neighbor-algorithm-in-Python-using-Sklearn-Pandas-and-NumPy/blob/main/Categorical/Categorical-Complete_case.py)

To impute Mixed datasets, run

5. [Mixed.py](https://github.com/Dalersingh-rs/Data-Imputation-using-k-nearest-neighbor-algorithm-in-Python-using-Sklearn-Pandas-and-NumPy/blob/main/Mixed/Mixed.py)

Based on the type of dataset to be imputed, the python file can be selected.

### Thank you :)
