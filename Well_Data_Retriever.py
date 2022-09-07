import os
# import numpy as np
import pickle
import pandas as pd
# from numba import jit, cuda

# @cuda.jit
def execution(wDir, **kwargs):
    counties = ['ADAMS', 'BARNES', 'BENSON', 'BILLINGS', 'BOTTINEAU', 'BOWMAN', 'BURKE',
    'BURLEIGH', 'CASS', 'CAVALIER', 'DICKEY', 'DIVIDE', 'DUNN', 'EDDY', 'EMMONS', 'FOSTER',
    'GOLDEN', 'GRAND', 'GRANT', 'GRIGGS', 'HETTINGER', 'KIDDER', 'LAMOURE', 'LOGAN', 'MCHENRY',
    'MCINTOSH', 'MCKENZIE', 'MCLEAN', 'MERCER', 'MORTON', 'MOUNTRAIL', 'NELSON', 'OLIVER',
    'PEMBINA', 'PIERCE', 'RAMSEY', 'RANSOM', 'RENVILLE', 'RICHLAND', 'ROLETTE', 'SARGENT',
    'SHERIDAN', 'SIOUX', 'SLOPE', 'STARK', 'STEELE', 'STUTSMAN', 'TOWNER', 'WALSH', 'WARD',
    'WELLS', 'WILLIAMS']

    os.chdir(wDir, )

    countyProdWells = {}

    wellsProd = {}
    poolProd = {}
    dateProd = {}

    for county in counties:
        countyProdWells[county] = []

        wellsProd[county] = {}
        poolProd[county] = {}
        dateProd[county] = {}

        for folderPath, wells, files in os.walk(os.path.join(wDir,county)):

            if files:
                well = folderPath[folderPath.index(county+'\\')+len(county)+1:]
                countyProdWells[county].append(well)
                try:
                    prodCSV = os.path.join(folderPath,'Production Data.csv')
                    prodData = pd.read_csv(prodCSV)
                except:
                    if os.stat(prodCSV).st_size == 0:
                        os.remove(prodCSV)
                    continue
                oil = sum(prodData.loc[:,'BBLS Oil'])
                gas = sum(prodData.loc[:,'MCF Prod'])
                water = sum(prodData.loc[:,'BBLS Water'])

                wellsProd[county][well] = {'Oil':oil,'Gas':gas,'Water':water} # Data for heat map 1 and 2
                
                for i in range(len(prodData)):
                    pool, date, oil, gas, water = prodData.loc[i,['Pool','Date','BBLS Oil','MCF Prod','BBLS Water']]

                    try:
                        poolProd[county][pool]['Oil'] += oil
                        poolProd[county][pool]['Gas'] += gas
                        poolProd[county][pool]['Water'] += water
                    except KeyError:
                        poolProd[county][pool] = {'Oil':oil,'Gas':gas,'Water':water}

                    try: # Data for 3
                        dateProd[county][date]['Oil'] += oil
                        dateProd[county][date]['Gas'] += gas
                        dateProd[county][date]['Water'] += water
                    except KeyError:
                        dateProd[county][date] = {'Oil':oil,'Gas':gas,'Water':water}
                print('Done: ',well,', ',county)

    return countyProdWells, wellsProd, poolProd, dateProd

if __name__ == '__main__':
    wDir = 'D:\\UND\\Data collection\\Well Data'
    countyProdWells, wellsProd, poolProd, dateProd = execution(wDir)
    
    with open(os.path.join(wDir,'Production Data.pkl'), 'wb') as f:  # Python 3: open(..., 'wb')
        pickle.dump([countyProdWells, wellsProd, poolProd, dateProd], f)