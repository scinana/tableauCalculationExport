# tableauCalculationExport

## What this code does
- This code will extract all Calculated Fields, Default Fields and Parameters from a Tableau workbook and export them into an Excel and PDF file.
- The code will also generate a Mermaid diagram showing the Lineage between fields. The diagram will be exported into an html file that you can open on your PCs Internet browser.
- Note that the Lineage Diagram will only show relationships between USED fields (ie. Default (datasource) fields that are NOT used in an Calculated Field will NOT come up in the diagram).

## Limitations and Important Considerations
- The latest version of the code will only work on **twbx** files (packaged Tableau files).
- The code is only available for **Windows** systems (as it needs the win32com.client package to generate the Excel file).

## Getting Started
- Please make sure you have a **working Python environment**, and you have installed the following packages/libraries (either via pip install or conda install - please Google the steps to install each package as some are either pip or Conda specific)
  - win32com.client
  - [tableaudocumentapi](https://tableau.github.io/document-api-python/docs/)
  - pandas
  - Jupyter Notebook
 - Some modules should already come with your Python installation (depending on what Python version you are using), but if for some reason they're not present in your Python env, please make sure you get them too
   - pathlib
  
## Downloading the Code and Setting up your working directory
1. Download the code into your preferred directory (ie. a folder on your PC).
  
 


