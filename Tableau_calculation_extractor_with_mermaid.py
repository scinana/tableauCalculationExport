# version 3.0

import pandas as pd, os, re, string, webbrowser

from tableaudocumentapi import Workbook
from os.path import isfile, join

import excelgenerator as exg

pd.set_option('display.max_columns', None)


# ## File Handling
# 
# - this version of code will only work with twbx files

input_path = "inputs"
output_path = "outputs"

mypath = "./{}".format(input_path)   #./ points to "this path" as a relative path

#only gets files and not directories within the inputs folder -https://stackoverflow.com/questions/3207219/how-do-i-list-all-files-of-a-directory
input_files = [f for f in os.listdir(mypath) if isfile(join(mypath, f))] 
input_files



def removeSpecialCharFromStr(spstring):
    
#     """
#     input: string
#     output: new string, without any special char
#     """
    
    return ''.join(e for e in spstring if e.isalnum())


def removeSpecialCharFromStr_leaveSpaces(spstring):
  
    return ''.join(e for e in spstring if (e.isalnum() or e ==' '))


def remove_sp_char_then_turn_spaces_into_underscore(string_to_convert):
    filtered_string = re.sub(r'[^a-zA-Z0-9\s_]', '', string_to_convert).replace(' ', "_")
    return filtered_string


def remove_sp_char_leave_undescore_square_brackets(string_to_convert):
    filtered_string = re.sub(r'[^a-zA-Z0-9\s_\[\]]', '', string_to_convert).replace(' ', "_")
    return filtered_string

def find_twbx_file(inputfile):
    
#     """
#     input: any input file
#     output: returns the file name without any special char for a twxb file if one is found, else returns empty string
#     """

    if inputfile[-5:] == '.twbx':
        sp_packagedWorkbook = i[:len(inputfile)-5]
       
        packagedWorkbook = removeSpecialCharFromStr(sp_packagedWorkbook)+'.twbx'
        
        old_file = join(input_path, sp_packagedWorkbook+'.twbx')
        new_file = join(input_path, packagedWorkbook)
        os.rename(old_file, new_file)

    else:
        packagedWorkbook = "" 
    
    return packagedWorkbook


for i in input_files:
    packagedWorkbook = find_twbx_file(i)
    print('Packaged workbook (no sp char): ' + packagedWorkbook)

    #substring to be used when naming the exported data, NEEDS A PACKAGED WORKBOOK TO EXIST, OTHERWISE IT WILL GIVE AN EMPTY STRING
    tableau_name_substring = packagedWorkbook.replace(".twbx","")[:30]
    print('\nOutput docs name (word/pdf): ' + tableau_name_substring)
    
packagedTableauFile_relPath = input_path+"/"+packagedWorkbook


# # Doc API

# get all fields in workbook
TWBX_Workbook = Workbook(packagedTableauFile_relPath)

collator = []
calcID = []
calcID2 = []
calcNames = []

c = 0

for datasource in TWBX_Workbook.datasources:
    datasource_name = datasource.name
    datasource_caption = datasource.caption if datasource.caption else datasource_name

    for count, field in enumerate(datasource.fields.values()):
        dict_temp = {
            'counter': c,
            'datasource_name': datasource_name,
            'datasource_caption': datasource_caption,
            'alias': field.alias,
            'field_calculation': field.calculation,
            'field_calculation_bk': field.calculation,
            'field_caption': field.caption,
            'field_datatype': field.datatype,
            'field_def_agg': field.default_aggregation,
            'field_desc': field.description,
            'field_hidden': field.hidden,
            'field_id': field.id,
            'field_is_nominal': field.is_nominal,
            'field_is_ordinal': field.is_ordinal,
            'field_is_quantitative': field.is_quantitative,
            'field_name': field.name,
            'field_role': field.role,
            'field_type': field.type,
            'field_worksheets': field.worksheets,
            'field_WHOLE': field
        }

        if field.calculation is not None:
            calcID.append(field.id)
            calcNames.append(field.name)

            f2 = field.id.replace(']', '').replace('[', '')
            calcID2.append(f2)

        c += 1
        collator.append(dict_temp)



def default_to_friendly_names2(formulaList,fieldToConvert, dictToUse):

    for i in formulaList:
        for tableauName, friendlyName in dictToUse.items():
            try:
                i[fieldToConvert] = (i[fieldToConvert]).replace(tableauName, friendlyName)
            except:
                a = 0
       
    return formulaList


def category_field_type(row):
    if row['datasource_name'] == 'Parameters':
        val = 'Parameters'
    elif row['field_calculation'] == None:
        val = 'Default_Field'
    else:
        val = 'Calculated_Field'
    return val

def compare_fields(row):
    if row['field_id'] == row['field_id2']:
        val = 0
    else:
        val = 1
    return val


calcDict = dict(zip(calcID, calcNames))
calcDict2 = dict(zip(calcID2, calcNames)) #raw fields without any []

collator = default_to_friendly_names2(collator,'field_calculation',calcDict2)

df_API_all = pd.DataFrame(collator)
df_API_all['field_type'] = df_API_all.apply(category_field_type, axis=1)

preference_list=['Parameters', 'Calculated_Field', 'Default_Field']
df_API_all["field_type"] = pd.Categorical(df_API_all["field_type"], categories=preference_list, ordered=True)

#get rid of duplicates for parameters, so only parameters from the explicit Parameters datasource are kept (as they are also listed again under the name of any other datasources)
df_API_all = df_API_all.sort_values(["field_id","field_type"]).drop_duplicates(["field_id", 'field_calculation']) 

df_API_all['field_id2'] = df_API_all['field_id'].str.replace(r'[\[\]]', '', regex=True)

df_API_all['comparison'] = df_API_all.apply(compare_fields, axis=1)
df_API_all = df_API_all[df_API_all['comparison'] == 1]

df_API_all = df_API_all.drop(['field_id2', 'comparison'], axis=1)
df_API_all.sort_values(['datasource_name', 'field_type', 'counter', 'field_name'])

df1 = df_API_all[[ 'field_name', 'field_datatype','field_type',  'field_calculation',   'field_id', 'datasource_caption']].copy()

preference_list=[ 'Default_Field', 'Parameters', 'Calculated_Field']
df1["field_type"] = pd.Categorical(df1["field_type"], categories=preference_list, ordered=True)
df1 = df1.sort_values(['field_type'])

df1.columns = ['Field_Name', 'DataType', 'Type', 'Calculation', 'Field_ID', 'Datasource']

df1['Field_Name'] = df1['Field_Name'].str.replace(r'[\[\]]', '', regex=True)



# ## Generating an excel file from a df (so the excel rows/cols can be formatted), then turning the excel into a pdf

#modify this part if you want to add more information/dfs to be saved as a separate sheet in excel

dfs_to_use = [{'excelSheetTitle': 'All fields extracted from DOC API', 'df_to_use':df1, 'mainColWidth':'' , 
               'normalColWidth': [10,15,50,20, 25], 'sheetName': 'GeneralDetails', 'footer': 'Data_1 (DOC API)', 'papersize':9, 'color': '#fff0b3'}                
             
             ]

#papersize: a3 = 8, a4 = 9



path_excel_file_to_create, path_pdf_file_to_create = exg.create_new_file_paths(tableau_name_substring+'_CALCS_only')

exg.create_excel_from_dfs(dfs_to_use, path_excel_file_to_create)

exg.create_pdf_from_excel(path_excel_file_to_create, path_pdf_file_to_create, dfs_to_use)


# # Start of mermaid module

# In[23]:


def first_char_checker(cell_value):
    if cell_value[0] != '[':
        cell_value = '__' + cell_value + '__'
    else:
        cell_value = cell_value.replace('[', '__')
        cell_value = cell_value.replace(']', '__')

    return cell_value


#define abc list to use during mermaid creation

abc=list(string.ascii_uppercase)
collated_abc = []

for i in abc:
    for j in abc:
        collated_abc.append(i+j)


# In[24]:


def_fields = df1[df1['Type'] == 'Default_Field']['Field_ID'].copy().apply(remove_sp_char_leave_undescore_square_brackets)

abc_touse = collated_abc[0:len(def_fields)]

def_fields_final = pd.DataFrame(list(zip(def_fields.tolist(), abc_touse)))
def_fields_final['aa'] = def_fields_final.apply(lambda row: first_char_checker(row[0]), axis=1)
def_fields_final['default_field'] = def_fields_final.apply(lambda row: '_st_' + row['aa'] + '_en_', axis=1)

mapping_dict_friendly_names = dict(zip(def_fields_final[0].tolist(), abc_touse))
mapping_dict = dict(zip(def_fields_final['aa'].tolist(), abc_touse))


created_calc = df_API_all[df_API_all['field_type'] != 'Default_Field'][['field_name', 'field_id', 'field_calculation', 'field_calculation_bk']].copy()

nlsi = ['x___' + i for i in collated_abc]
nlsi_to_use = nlsi[0:len(created_calc)]

created_calc['field_name'] = created_calc['field_name'].apply(remove_sp_char_leave_undescore_square_brackets)
created_calc['aa'] = created_calc.apply(lambda row: first_char_checker(row['field_id']), axis=1)
created_calc['calc_field'] = created_calc.apply(lambda row: '_st_' + row['aa'] + '_en_', axis=1)
created_calc['field_calculation_bk'] = created_calc['field_calculation_bk'].str.replace(r'[\[\]]', '__', regex=True)

calc_map_dict_friendly_names = dict(zip(created_calc['field_name'].to_list(), nlsi_to_use))
calc_map_dict = dict(zip(created_calc['aa'].to_list(), nlsi_to_use))


def create_mermaid_paths(df, field_type):
    
    c = 0
    t_collator = []

    for i in df['aa']:

        print('\n______________________' + field_type.upper() + ' TO ANALYSE ________________________: ' + i + '\n')

        try:
            tlist = created_calc[created_calc['field_calculation_bk'].str.contains(i, regex=False) == True]['aa'].to_list()
        except:
            tlist = []

        if len(tlist) != 0:
            print('LIST PRINTING:\n\n' + str(tlist))

            for x in tlist:
                newdict = {}

                newdict['count'] = c
                newdict['starting'] = i
                newdict['ending'] = x

                newdict['path_mermaid'] = i + " --> " + x

                print('\n' + str(c) + ' ******************NEW DICT PRINTING ********************** \n\n' + str(newdict))

                t_collator.append(newdict)

                c = c + 1
    
    return t_collator



t_collator_def_fields = create_mermaid_paths(def_fields_final, 'default_field')


t_collator_calcs = create_mermaid_paths(created_calc, 'calculation')



###############################
#replace the full names of fields and calcs for their abbrv letters, to make the mermaid code leaner

for default_field, mapping_letter in mapping_dict.items():
    for i in t_collator_def_fields:
        i['path_mermaid'] = i['path_mermaid'].replace(default_field, mapping_letter)

for default_field, mapping_letter in calc_map_dict.items():
    for i in t_collator_def_fields:
        i['path_mermaid'] = i['path_mermaid'].replace(default_field, mapping_letter)


##############################

##############################
# replace the full names of fields and calcs for their abbrv letters, to make the mermaid code leaner

for default_field, mapping_letter in mapping_dict.items():
    for i in t_collator_calcs:
        i['path_mermaid'] = i['path_mermaid'].replace(default_field, mapping_letter)

for default_field, mapping_letter in calc_map_dict.items():
    for i in t_collator_calcs:
        i['path_mermaid'] = i['path_mermaid'].replace(default_field, mapping_letter)


##############################


new_list_a = ['']
fields_list = ['']

new_list_a.extend([i['path_mermaid'] for i in t_collator_calcs])
new_list_a.extend([i['path_mermaid'] for i in t_collator_def_fields])

################################
#find the unique nodes within the a --> b mermaid paths in new_list_a (eg. a and b)
c = []

for i in new_list_a:
    print(i)
    c.append(i.split(' --> ')[0])

    try:
        c.append(i.split(' --> ')[1])
    except:
        pass

c.pop(0)
s = set(c)
c = list(s)
##############################

for i, d in mapping_dict_friendly_names.items():
    if d in c:
        if i[0] != '[':
            print(d + "[" + i + "]")
            fields_list.append(d + "[" + i + "]:::foo")
        else:
            print(d + i)
            fields_list.append(d + i + ':::foo')

for i, d in calc_map_dict_friendly_names.items():
    if d in c:
        print(d + "[" + i + "]")
        fields_list.append(d + "[" + i + "]")
        
superfinallist =  fields_list + new_list_a



mermaid_diagram_code = """
flowchart LR
    classDef foo fill:#f9f,stroke:#333,stroke-width:1px{}
""".format("\n\t".join(superfinallist))

print(mermaid_diagram_code)


### Create html which will display the mermaid diagram

html_base = """

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>""" + tableau_name_substring + " Calculation Lineage" + """</title>
    <!-- Include Mermaid.js library -->
    <script type="module">
      import mermaid from 'https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.esm.min.mjs';
      mermaid.initialize({ startOnLoad: true });
    </script>
</head>
<body>
    <h1>""" +  tableau_name_substring + " Calculation Lineage" + """</h1>
    <!-- Mermaid diagram definition -->
    <div class="mermaid">""" + mermaid_diagram_code + """</div>
</body>
</html>
"""

print('\n ______________________________ START_OF_HTML ______________________________')
print(html_base)
print('\n ______________________________ END_OF_HTML ______________________________')



### Output html string to a local file, then open it on the web browser (this bit was done with help of chatgpt)

# Specify the file path
file_path = 'outputs\mermaid_diagram_{}.html'.format(tableau_name_substring)

# Write the string to an HTML file
with open(file_path, 'w') as file:
    file.write(html_base)

print("HTML content successfully written to {}".format(file_path))

# Open the HTML file in the default web browser
webbrowser.open('file://' + os.path.realpath(file_path))

### end of code block done with help of chatgpt

