#import all necessary packages
import numpy as np
import pandas as pd
import math
import os
import glob

#import KNNImputer from sklearn.impute
from sklearn.impute import KNNImputer

#read the NRMS/AE excel file and convert it into a numpy array
nrms=pd.read_excel(r'C:\Users\RANJEETH SINGH\Desktop\Table-NRMS-AE.xlsx')
nrmsArray=nrms.to_numpy()

#initialize count
count=0

#set path=file_name in which the excel files are present
path = r'D:\Daler\UNIVERSITY OF WINDSOR\Semester 3\Data mining\Project\Dataset and instructions\Incomplete Datasets\Incomplete Datasets Without Labels\Spam'

#to read all the excel files present, use glob. Doing so, all excel files are loaded in csv_files
csv_files = glob.glob(os.path.join(path, "*.xlsx"))

#for loop to run over all the excel files
for f in csv_files:
    # load the original dataset
    real = pd. read_excel(r'D:\Daler\UNIVERSITY OF WINDSOR\Semester 3\Data mining\Project\Dataset and instructions\Original Datasets\Original Datasets Without Labels\Spam.xlsx', header=None)

    # load the incomplete dataset and copy the content
    df = pd.read_excel(f, header=None)
    data = df.copy()

    #complete dataframe
    df2 = pd.DataFrame(df.dropna())

    #Missing dataframe
    Incomplete_data = df[df.isnull().any(axis=1)]
    df1 = pd.DataFrame(Incomplete_data)

    # Loop runs only for the Missing values and calculates only with the complete instances
    for i in range(len(df1)):
        first = df1.head(1)
        df2 = pd.concat([first, df2], axis=0, ignore_index=False)

        # Removing the first row from the Missing Datafile
        df1 = df1.iloc[1:, :]

        # kNN imputation starts here for complete case
        k = int(math.sqrt(len(df2)))
        imputer = KNNImputer(n_neighbors=k)
        After_imputation = imputer.fit_transform(df2.sort_index(axis=0))

    # compute NRMS
    nrms_numerator = np.linalg.norm(After_imputation - real)
    nrms_denomerator = np.linalg.norm(real)
    nrms = nrms_numerator / nrms_denomerator

    # splitting the path to generate the excel file name from it
    split_f=(os.path.split(f))
    dataset_name=split_f[1].split(".")

    # find the same excel file name in te numpy nrms_array and insert the nrms value
    wh = np.where(nrmsArray ==dataset_name[0])
    r = wh[0][0]
    c = wh[1][0]
    nrmsArray[r][c + 6] = nrms
    nrmsArray[r][c + 8] = k

    # use count to keep track of the number of datasets that have been imputed in the loop
    count+=1
    print(count)

# paste the numpy Array in the NRMS/AE Excel file
pd.DataFrame(nrmsArray).to_excel(r'C:\Users\RANJEETH SINGH\Desktop\Table-NRMS-AE.xlsx',index=False)