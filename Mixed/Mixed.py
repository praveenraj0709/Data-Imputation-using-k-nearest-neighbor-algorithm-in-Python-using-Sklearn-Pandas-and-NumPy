# import all necessary packages
import numpy as np
import pandas as pd
import math
import os
import glob

# import KNNImputer from sklearn.impute
from sklearn.impute import KNNImputer
from sklearn.preprocessing import OrdinalEncoder
from numpy import asarray

# read the NRMS/AE excel file and convert it into a numpy array
nrms = pd.read_excel(r'C:\Users\anish\OneDrive\Desktop\Table-NRMS-AE.xlsx')
nrmsArray = nrms.to_numpy()

# initialize count
count = 0

# set path=file_name in which the excel files are present
path = r'C:\Users\anish\Downloads\Incomplete Datasets\Incomplete Datasets Without Labels\Abalone'

# to read all the excel files present, use glob. Doing so, all excel files are loaded in csv_files
csv_files = glob.glob(os.path.join(path, "*.xlsx"))

# for loop to run over all the excel files
for f in csv_files:
    # load the original dataset
    real = pd.read_excel(r'C:\Users\anish\Downloads\Original Datasets\Original Datasets Without Labels\Abalone.xlsx',
                         header=None)

    # load the incomplete dataset and copy the content
    df = pd.read_excel(f, header=None)
    data = df.copy()

    # set value of k (is usually square root of the total number of instances)
    k = int(math.sqrt(len(df)))


    real_num = real.select_dtypes(include=[np.number])
    real_cat = real.select_dtypes(exclude=[np.number])
    numeric_data = df.select_dtypes(include=[np.number])
    categorical_data = df.select_dtypes(exclude=[np.number])

    # define ordinal encoding
    encoder = OrdinalEncoder()

    real_result = encoder.fit_transform(real_cat)
    # real_result.columns=real_cat.columns
    real_enc_new = pd.DataFrame(real_result)

    # transform data
    result = encoder.fit_transform(categorical_data)
    # result.columns=categorical_data.columns
    X_enc_new = pd.DataFrame(result)
    # X_enc_new.columns=categorical_data.columns
    # print(X_enc_new)

    from sklearn.impute import KNNImputer

    imputer = KNNImputer(n_neighbors=k)
    cat_imputation = imputer.fit_transform(result)
    num_imputation = imputer.fit_transform(numeric_data)
    cat_imputation.round(0)
    num_imputation.round(0)

#To crosscheck the imputed values, we have transferred the values to CSV, but here we have commented those 2 lines
    # pd.DataFrame(cat_imputation).to_csv('imputed_cat.csv', index=None)
    # pd.DataFrame(num_imputation).to_csv('imputed_num.csv', index=None)

    # print("Categorial")
    reverse_data_x = encoder.inverse_transform(cat_imputation)
    final = pd.DataFrame(reverse_data_x)

    # compute nrms
    nrms_numerator = np.linalg.norm(num_imputation - real_num)
    nrms_denomerator = np.linalg.norm(real_num)
    nrms = nrms_numerator / nrms_denomerator
    print(nrms)

    # compute AE
    comp = (real_cat.to_numpy() == final.to_numpy())
    comp = pd.DataFrame(comp).replace({True: 1, False: 0})
    sumOfV = comp.values.sum()
    total = comp.count().sum()
    sumOfV, total
    AE = round(sumOfV / total, 4)
    print(AE)

print("COMPLETED")
