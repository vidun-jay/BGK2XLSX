from Bio import SeqIO
import pandas as pd

# Function to extract gene information from GenBank file
def extract_genes_from_genbank(file_path):
    gene_data = []
    for record in SeqIO.parse(file_path, "genbank"):
        for feature in record.features:
            if feature.type == "gene":
                gene_info = {
                    "Gene": feature.qualifiers.get('gene', [''])[0],
                    "Locus_Tag": feature.qualifiers.get('locus_tag', [''])[0],
                    "Product": feature.qualifiers.get('product', [''])[0]
                }
                gene_data.append(gene_info)
    # Create a DataFrame from the gene data
    df = pd.DataFrame(gene_data)
    return df

# Function to save DataFrame to Excel
def save_to_excel(df, output_file):
    df.to_excel(output_file, index=False, engine='openpyxl')

# prompt for input path
genbank_file = input("Drag and drop your GenBank file to the terminal window: ")
output_file = 'gene_data.xlsx'  # output file in the current directory

# Extract genes
gene_data = extract_genes_from_genbank(genbank_file)

# Save the data to Excel
save_to_excel(gene_data, output_file)

print(f"\033[92mGene data has been saved to {output_file}\033[0m")

input("Press Enter to exit...")