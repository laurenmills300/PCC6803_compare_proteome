import os
import sys
import multiprocessing
import subprocess
from subprocess import check_output
from multiprocessing import process
import urllib.request
import pandas as pd
import numpy as np
from Bio.Blast.Applications import NcbiblastpCommandline
global data_url
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
import time
import numpy
import csv
from Bio import Entrez
import re

accession = open('input_files\Syn11901.faa', 'r')

# all WP_ accessions in the PCC 11901 genome
accession_dict = {}
name = "none"
sequence = []
wp_genome = []
for line in accession:
    if line.startswith('>WP_'):
        if name == "none":
            pass
        else:
            #print(name, "".join(sequence))
            accession_dict[accession] = name + "\t" + (str(len("".join(sequence))))
        data = line.split(' ')
        data = data[0].split('>')
        #print(data)
        wp_genome.append(str(data[1]))
        #wp.split('>')
        #accession = line.split(' ')[0].strip('>').strip()
        #print(accession)
        #name = name.split(' [')
        #name = name[0]
        #print(name)

        sequence = []
        #print(name)
    else:
        sequence.append(line.strip())
df_wp_genome = pd.DataFrame({'WP': wp_genome})


# WP from blast hits against PCC 6803
df_blast = pd.read_excel('input_files/input_unique_proteins_all_6803_proteins.xlsx')

#df_blast.loc[df_blast[1].startswith('WP_')]

df_wp_blast = df_blast.loc[df_blast[1].str.contains('WP_')]
print(df_wp_blast)
df_wp_blast.reset_index(drop=True, inplace=True)
wp_blast = df_wp_blast[1]


wp_blast_no_duplicates = wp_blast.drop_duplicates()
wp_blast_no_duplicates = wp_blast_no_duplicates.values.tolist()

#turn into sets and find unique values
wp_blast_no_duplicates = set(wp_blast_no_duplicates)
wp_genome = set(wp_genome)

unique_set = set(wp_genome)-set(wp_blast_no_duplicates)
unique_list = list(unique_set)
print(len(unique_list))
uniqueprotein_textfile = open('output_files/unique_proteins_all.txt', 'w')
uniqueprotein_textfile.write(str(unique_list))
uniqueprotein_textfile.close

# format the dataframe for information on unique list
#int_dataframe = pd.read_table('output_files/unique_proteins_all.txt')
#int_dataframe.to_excel('output_files/unique_proteins_all.xlsx', 'Sheet 1')

int_df = pd.DataFrame(columns=['NCBI Accession'])
for accession in unique_list:
    temporary_df = pd.DataFrame([accession], columns=['NCBI Accession'])
    int_df = int_df.append(temporary_df, ignore_index=True)
print(int_df)

dataframe = int_df
dataframe = dataframe.iloc[:, :2]


def search_data(ncbiTaxId):

    Entrez.email = 'INSERT YOUR EMAIL HERE'
    searchResultHandle = Entrez.esearch(db="protein", term=ncbiTaxId)
    searchResult = Entrez.read(searchResultHandle)

    id = searchResult["IdList"]
    searchResultHandle.close()

    entryData = Entrez.efetch(db="protein", id=id, rettype='gp').read()

    return entryData

def parse(dataset: str):
    print(dataset)

    try:
        name_raw = re.findall(r'product=".*"', dataset)
        name_raw = name_raw[0]
    except:
        name_raw = re.findall(r'product=".*\n.*"', dataset)
        name_raw = name_raw[0]
        name_raw = " ".join(name_raw.split())
    print(f'Raw name: {name_raw}')

    name_raw = name_raw.split('=')
    name = name_raw[-1]
    name = name.replace('"', '')

    length_raw = re.search("[0-9]*.aa", dataset).group()
    length_raw = length_raw.split(' ')
    length = length_raw[0]

    try:
        mol_weight_raw = re.search('calculated_mol_wt=[0-9]*', dataset).group()
        mol_weight_raw = mol_weight_raw.split('=')
        mol_weight = mol_weight_raw[-1]
    except Exception as e:
        print(e)
        mol_weight = None


    date_raw = re.search('BCT.[0-9]*-[A-Z]*-[0-9]*', dataset).group()
    date_raw = date_raw.split(' ')
    date = date_raw[-1]

    return name, length, mol_weight, date

NCBI_IDs = dataframe.iloc[:, 0]

ID_name = []
ID_length = []
ID_mol_weight = []
ID_date = []

counter = 0
for ID in NCBI_IDs:

    print(f"\n\nLooking at {ID}...")
    name, length, mol_weight, date = parse(search_data(ID))
    ID_name.append(name)
    ID_length.append(length)
    ID_mol_weight.append(mol_weight)
    ID_date.append(date)
    print(f"Appended data:")
    print(f"Name: {name}, Length: {length}, Molecular Weight: {mol_weight}, Date created: {date}")
    counter += 1

dataframe['Name'] = ID_name
dataframe['Length (aa)'] = ID_length
dataframe['Calculated mol wt'] = ID_mol_weight
dataframe['Date Added to NCBI'] = ID_date

print(f"Final Dataframe:\n {dataframe}")

wb = Workbook()
global results_file
timestr = time.strftime('%Y_%m_%d-%H_%M')
results_file = 'output_files\{}_unique_protein_information.xlsx'.format(timestr)
dataframe.to_excel(results_file, index=False)
