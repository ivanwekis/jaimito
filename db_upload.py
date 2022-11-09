from queue import Empty
#from pymongo import MongoClient
import argparse
import certifi
import pandas as pd 

BASE_PATH = "C:/Users/jaime.sanchez/scripts/"


def read_csv(input_file):
    path = BASE_PATH.join(input_file)
    data = pd.read_excel(path ,parse_dates=True)
    data.info
    return data

def main(input_file):
    FILE_NAMES = ["Estado_de_transportes_externos_SEMAT,_S.A._Autohero_outbound.xlsx","Estado_de_transportes_externos_SEMAT,_S.A._Autohero_inbound.xlsx","Estado_de_transportes_externos_SEMAT_S.A._Remarketing.xlsx","Estado_de_transportes_externos_SEMAT_S.A._Auto1.xlsx"]
    
    for file in FILE_NAMES:
        data_list_dictionary = read_csv(file)
        #write_csv(input_file,data_list_dictionary)


if __name__ == "__main__":
    main()