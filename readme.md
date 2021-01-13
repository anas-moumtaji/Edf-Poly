# Code For Word Report Automation Using pydoc


## Pandas and Pydoc

This section of code show how to read data from Excel File and write it directly into the Word Document

```python
excel_data_df = pandas.read_excel('Ressources/synthese_comparative_FigSeq.xlsx', sheet_name='1_PI_existant HT_evolutions HT', usecols = "A:C",skiprows = range(11, 58))
print(excel_data_df)

t = document.add_table(excel_data_df.shape[0] + 1, excel_data_df.shape[1])
t.style = 'Light List Accent 2'

for j in range(excel_data_df.shape[-1]):
    if str(excel_data_df.columns[j]) == "Unnamed: 0":
        excel_data_df.columns[j] == ""
    else:
        t.cell(0, j).text = excel_data_df.columns[j]

# add the rest of the data frame
for i in range(excel_data_df.shape[0]):
    for j in range(excel_data_df.shape[-1]):
        if str(excel_data_df.values[i, j]) == "nan":
            t.cell(i + 1, j).text == ""
        else:
            t.cell(i + 1, j).text = str(excel_data_df.values[i, j])


```
