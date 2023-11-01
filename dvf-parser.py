import re;
import argparse;
import numpy as np;
import pandas as pd;
from styleframe import StyleFrame;

fieldsLabels = {
    'road': 'Rue',
    'city': 'Ville',
    'price': 'Prix',
    'surface2': 'Surface',
    'roomsCount': 'Pieces',
    'cityName': 'nom_commune',
    'value': 'valeur_fonciere',
    'cityCode': 'code_commune',
    'cityCode2': 'Code commune',
    'priceSquareMeter': 'Prix au m2',
    'surface': 'surface_reelle_bati',
    'addressNumber': 'adresse_numero',
    'addressName': 'adresse_nom_voie',
    'rooms': 'nombre_pieces_principales',
}


def formatNumber(stringNumber):
    return format(stringNumber, ',').replace(',', ' ');

def readFile(filePath):
    dataframe = pd.read_excel(filePath);
    # dataframe = dataframe[dataframe["adresse_nom_voie"].str.contains("None") == False] 
    dataframe[fieldsLabels['value']] = dataframe[fieldsLabels['value']].fillna(0);
    dataframe[fieldsLabels['surface']] = dataframe[fieldsLabels['surface']].fillna(1);
    return dataframe;

def gnerateNewDataframe(fileDataFrame):
    areas = [];
    roads = [];
    rooms = [];
    cityCodes = [];
    cityNames = [];
    sellPrices = [];
    pricesPerSquareMeteres = [];
    newDataframe = pd.DataFrame();

    for idx in fileDataFrame.index:
        cityNames.append(fileDataFrame[fieldsLabels['cityName']].iloc[idx]);
        cityCodes.append(fileDataFrame[fieldsLabels['cityCode']].iloc[idx]);
        areas.append(formatNumber(round(fileDataFrame[fieldsLabels['surface']].iloc[idx])));
        sellPrices.append(formatNumber(round(fileDataFrame[fieldsLabels['value']].iloc[idx])));
        rooms.append(fileDataFrame[fieldsLabels['rooms']].iloc[idx]);
        roads.append(f"{str(fileDataFrame[fieldsLabels['addressNumber']].iloc[idx]).replace('.0', '')} {fileDataFrame[fieldsLabels['addressName']].iloc[idx].lower()}");

        # Sensitive condition 
        if fileDataFrame[fieldsLabels['addressName']].iloc[idx] == 'None' and idx > 1: 
            roads[-1] = 'NaN';

        if fileDataFrame[fieldsLabels['surface']].iloc[idx] == 1:
            pricesPerSquareMeteres.append('NaN');
        else:
            pricesPerSquareMeteres.append(
                    formatNumber(
                        round(
                            fileDataFrame[fieldsLabels['value']].iloc[idx] / fileDataFrame[fieldsLabels['surface']].iloc[idx]
                        )
                    )
                );


    newDataframe[fieldsLabels['cityCode2']] = cityCodes;
    newDataframe[fieldsLabels['city']] = cityNames;
    newDataframe[fieldsLabels['road']] = roads;
    newDataframe[fieldsLabels['surface2']] = areas;
    newDataframe[fieldsLabels['roomsCount']] = rooms;
    newDataframe[fieldsLabels['price']] = sellPrices;
    newDataframe[fieldsLabels['priceSquareMeter']] = pricesPerSquareMeteres;


    return newDataframe;

def generateXLSX(generatedDataFrame):
    excel_writer = StyleFrame.ExcelWriter('./dvf.xlsx');

    sf = StyleFrame(generatedDataFrame);
    sf.set_column_width(columns=[fieldsLabels['road']], width=45);
    sf.set_column_width(columns=[fieldsLabels['price']], width=20);
    sf.set_column_width(columns=[fieldsLabels['surface2'], fieldsLabels['roomsCount']], width=15);
    sf.set_column_width(columns=[fieldsLabels['cityCode2'], fieldsLabels['city'], fieldsLabels['priceSquareMeter']], width=20);
    sf.to_excel(excel_writer, row_to_add_filters=0);

    excel_writer.save();
    


def main ():
    argparser = argparse.ArgumentParser();
    argparser.add_argument('-f', '--file', help="Xlsx file path");

    args = argparser.parse_args();

    if args.file is None:
        print('Please specify the path for the file to parse');
        return;
    
    if args.file and ('.xlsx' not in args.file):
        print('Please select a valid xlsx file');
        return;

        
    fileDataFrame = readFile(args.file);
    generatedDataFrame = gnerateNewDataframe(fileDataFrame)
    generateXLSX(generatedDataFrame)



main();