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

**Before starting on this section, please make sure you've installed Python and any dependencies into your Python environment (ie. the libraries and packages detailed in the previous section)**

1. Download ALL the code into your preferred directory (ie. a folder on your PC).
  - Make sure the **excelgenerator.py** file is in the SAME directory as the ipynb or py file you want to run (ie. the Tableau_calculation_extractor_with_mermaid.ipynb)

2. In your working directory, create an empty "/inputs" and an "/outputs" folder. Your working directory should look like this:

   Note that an "/inputs" folder means that you will create a folder called "inputs" inside your working directory. From here onwards I will use "/inputs" and "inputs" interchangeably (same for outputs).

![image](https://github.com/user-attachments/assets/62ec66c6-0db6-495a-9063-8b603fe66d17)


 
3. Once you have a Tableau packaged workbook (twbx file) that you want to analyse, save it in the "inputs" folder.

4. Run your **Calculation Extractor code** (ie. Tableau_calculation_extractor_with_mermaid.ipynb or Tableau_calculation_extractor_with_mermaid.py, depending on which version you want to run - either a Jupyter Notebook one or a py file - they're both meant to have the same functionality)
   
5. Check the "/outputs" folder for the code outputs - you should now have a PDF, Excel and HTML file with the results from the Calculation Extraction process (PDF and Excel) and the Lineage Creation process (the HTML file).

### Running the code again (eg. to analyse a new workbook)
At the moment the code will only run on one twbx at a time, and will **only handle 1 twbx file from the inputs folder**. If two or more twbx files are found in the inputs folder, the code will only analyse one of them --> in future versions of the code, I will add file handling so more than one twxb file can be analysed at a time - you can also submit a PR with this code if you'd like to contribute to this code!

- Before analysing a new workbook (once already saved to the "/inputs" folder), remove any OTHER files from the "/inputs" folder (eg. any previous workbook you have already analysed), and only leave the one workbook you want to analyse.
- You can now run the Calculation Extractor code.
- You don't need to worry about emptying the "/outputs" folder - this folder will simply store all the outputs from any runs of the Calculation Extractor code, so more and more outputs will be added as more runs occur.


 # Troubleshooting and Help
 As this is a personal project, I am not providing any IT support for this code. However if you have any questions that are NOT explained above, feel free to reach out to nana7milana@gmail.com.
 I will aim to reply within one or two weeks, but if I don't, feel free to send me a reminder.
 Thanks for checking out my code!

 Ana
  
 


