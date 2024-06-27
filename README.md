This Program read Italian digital XML invoice, For the purpose it was conceived it read the data from a directory named XML and parse the following fields:

    'FatturaElettronicaHeader/CessionarioCommittente/DatiAnagrafici/Anagrafica/Denominazione' 
    'FatturaElettronicaHeader/CedentePrestatore/DatiAnagrafici/Anagrafica/Denominazione'
    'FatturaElettronicaHeader/CedentePrestatore/DatiAnagrafici/Anagrafica/Cognome'
    'FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/Data'
    'FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/ImportoTotaleDocumento'
    'FatturaElettronicaBody/Pagamento/DettaglioPagamento/ImportoPagamento'
    'FatturaElettronicaBody/DatiGenerali/DatiGeneraliDocumento/Numero'
    'FatturaElettronicaBody/DatiBeniServizi/DettaglioLinee/Descrizione'

Then it will compose a Fundflow excel to summarize that dataset.


