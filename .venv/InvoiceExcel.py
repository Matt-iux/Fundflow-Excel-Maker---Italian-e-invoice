import xml.etree.ElementTree as ET
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl import load_workbook
from os import listdir, path
"""
This program read a group of electronic invoice in a XML format inside his directory searching for specific fields.
It gets the value of those fields and proceed to compose a dataframe first and an excel file filling a determinated 
structure with the element of the dataframe composed.
"""

def main():
    # Parse the Xml files to  generate a dataframe of the required structure
    df = Compose_dataframe()
    # Transpose the dataframe into a file excel
    Write_Excel(df)
    # Set a colour for the header row
    colour_row()


def Compose_dataframe():
    # Assign directory path for the xml file type
    directory_xml = r'.\XML'
    #Define the list to append item parsed
    all_items=[]

    xml_files = [path.join(directory_xml, f) for f in listdir(directory_xml) if f.endswith('.xml')]

    # Cycle through xml files
    for xml_file in xml_files:
        print(xml_file)
        tree = ET.parse(xml_file)
        root = tree.getroot()

        # Extract required element from XML
        Debitor = tree.find('FatturaElettronicaHeader/CessionarioCommittente/DatiAnagrafici/Anagrafica/Denominazione') # Group Company
        Beneficiary = tree.find('FatturaElettronicaHeader/CedentePrestatore/DatiAnagrafici/Anagrafica/Denominazione')
        # Add a correction  for invoice from Name and surname origin with no denomination
        if Beneficiary is None:
            Beneficiary = tree.find('FatturaElettronicaHeader/CedentePrestatore/DatiAnagrafici/Anagrafica/Cognome')
        Date = tree.find('FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/Data')
        Amount_Requested = tree.find('FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/ImportoTotaleDocumento')
        Cash_Flow = tree.find('FatturaElettronicaBody/DatiPagamento/DettaglioPagamento/ImportoPagamento')
        Invoice = tree.find('FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/Numero')
        Invoice_String = Invoice.text
        print(Invoice.text)
        Description = tree.find('FatturaElettronicaBody/DatiBeniServizi/DettaglioLinee/Descrizione')

        # Compose the invoice excel field as FT + invoice number + del + date of invoice
        Invoice_Excel_Field = (' FT ' + Invoice_String + ' del ' + Date.text)
        print(Invoice_Excel_Field)

        # Temporary Bugfix, Some invoice don't have the
        # FatturaElettronicaBody/DatiPagamento/DettaglioPagamento/ImportoPagamento field
        if Cash_Flow is None:
            Cash_Flow_Text = str("verificare su fattura")
        else:
            Cash_Flow_Text = Cash_Flow.text

        # Temporary Bugfix, Some invoice don't have the
        # FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/ImportoTotaleDocumento field
        if Amount_Requested is None:
            Amount_Requested_Text = str("verificare su fattura")
        else:
            Amount_Requested_Text = Amount_Requested.text

        # Set the dataframe, need to insert data, defined only columns
        df = pd.DataFrame({ 'Debitor' : Debitor.text,
                            'Beneficiary' : Beneficiary.text,
                            'Transaction Date': None,
                            #'Amount Request' : float(Amount_Requested.text), old field pre bugfix 2
                            'Amount Request': Amount_Requested_Text,
                            #'Cash Flow' : float(Cash_Flow.text), old field pre bugfix 1
                            'Cash Flow' : Cash_Flow_Text,
                            'Invoice' : Invoice_Excel_Field,
                            'Description' : Description.text,
                            'Purpose': None,
                            'Authorized' : None,
                            'Payment' : None,
                            'Notes' : None
        }, index=[1])
        # Put the extracted items in a list, otherwise they will get overwritten
        all_items.append(df)

    # Remake the dataframe concatenating the items extracted and saved in the list
    df = pd.concat(all_items, ignore_index= True)


    #Print df to check
    print('Here it is a dataframe composition spoiler',df)
    return(df)

# Print Excel with defined Column width and set index to false
def Write_Excel(df):
    with pd.ExcelWriter("Fundflow_Generato.xlsx",) as writer:
        df.to_excel(writer,index=False)
        # Style the cell column width and color
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        worksheet.set_column(0, 0, 15)
        worksheet.set_column(1, 2, 30)
        worksheet.set_column(3, 4, 20)
        worksheet.set_column(5, 11, 30)

# Colour the first row
def colour_row():
    wb = openpyxl.load_workbook("Fundflow_Generato.xlsx")
    ws = wb['Sheet1']
    cell_ids = ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1', 'J1', 'K1']
    for i in range(11):
        ws[cell_ids[i]].fill = PatternFill(patternType='solid', fgColor='DCE6F1')
    wb.save("Fundflow_Generato.xlsx")


if __name__ == "__main__":
    main()