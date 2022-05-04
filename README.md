# Synechocystis sp. PCC 6803: compare proteome with newly discovered cyanobacterial species

Here, we describe a method of conducting an all vs. all protein comparison using this bioinfomratics tool for analysing the model cyanobacteria Synechocystis sp. PCC 6803 (PCC 6803) against the proteome of newly discovered cyanobacterial species. The analysis incorporates the use of the National Centre for Biotechnology Information’s (NCBI) basic local alignment search tool (BLAST).

In this example, we use the newly discovered, fast-growing strain Synechococcus sp. PCC 11901 for comparison with PCC 6803. 

To compare against the well-documented PCC 6803 species, the input data was based upon the findings from Baers et al. (2019). A modified version of Supplementary Table 3 (ST3)  from this paper was created and formed the starting point for all the input data required for PCC 6803’s proteome information. This can be found input_data/ST3_Baers-SuppTable3_all-PCC-6803-proteins.xlsx.

A second input file was then taken from either Supplementary Tables 1, 3 or 4 found in Mills et al., these three tables are split into different class of PCC 6803’s proteome: central metabolism and transport; characterised proteins not involved in central metabolism; uncharacterised proteins. These can all be found under the input_data folder. These can either be run separately or can be combined, depending onthe output required by the user. 


For the first step in the analysis, PCC6803_comparative_analysis.py must be run. In this example, the central metabolsim table (ST1) is used. The python code has detailed notes within the script to explain each step, but below are a few points which should be highlighted for users.

Outline of PCC6803_comparative_analysis.py
1. Lines 13 - 149, formats and concatenates the two original excel files into one dataframe required for the analysis.

2. Line 169 can be changed to allow for the proteome file of any cyanobacterial species to be used here. This must be downloaded by the user from the NCBI's website. 

3. Line 171 uses the subprocess function to create a local BLAST database for the cyanoabcterial strain. For this, the BLAST+ must be installed on the users computer. See https://blast.ncbi.nlm.nih.gov/Blast.cgi?CMD=Web&PAGE_TYPE=BlastDocs&DOC_TYPE=Download for more details.

4. Line 213 is where the X for the BLAST search can be changed, the output format can be changed, as well as the e-value cut off.

The code then goes on to format the output for an easy-to-read, colour cooded table showing the hits associated with each well documetned PCC 6803 protein. 

The second code, unique_proteins_information.py, can be used to identify any protein IDs that did not match any of PCC 6803's proteome and can therefore be described as distinctly different (depending on the E-value chosen). Again, this code is heavily annotated but a few points to note:

1. Line 21, can be changed by the used for the appropriate FASTA file for the species used in analsyis. 
2. Line 94, insert an email address into the Entrez.email line to increase the speed od the code. 

More information from the Entrez database may be available depending on the species used for the comparison. The code can easily be adjusted to add this extra information in.


References:
Baers LL, Breckels LM, Mills LA, Gatto L, Deery MJ, Stevens TJ, et al. Proteome Mapping of a Cyanobacterium Reveals Distinct Compartment Organization and Cell-Dispersed Metabolism. Plant Physiol. 2019 Dec 1;181(4):1721–38.

Mills LA, McCormick AJ, Lea-Smith DJ. Current knowledge and recent advances in understanding metabolism of the model cyanobacterium Synechocystis sp. PCC 6803. Biosci Rep. 2020 Apr 3; 40(BSR20193325).
