import pandas as pd
import json

class Excel_to_Json:
    def __init__(self, file_, tab, column_to_drop, output_file_name):
        self.file_ = file_
        self.tab = tab
        self.column_to_drop = column_to_drop
        self.output_file_name = output_file_name + '.json'

    def Convert(self):

        decision_table = pd.read_excel(self.file_, sheet_name = self.tab) 
        tb = pd.DataFrame(decision_table)
        filter_table = tb.drop(columns = self.column_to_drop)
        json_string = pd.DataFrame.to_json(filter_table, orient = 'index') 
        s = json.loads(json_string)

        with open(self.output_file_name,'w') as json_file:
            json.dump(s, json_file, indent=3) 
            
        print('Arquivo %s gerado com sucesso' %(self.output_file_name))

#Discret template conversion
discret_template = Excel_to_Json(file_='DeTableV77.xlsx', 
                                tab='Templates - Discretos', 
                                column_to_drop=['Descricao','BlockType','ElemName','Elem Type'],
                                output_file_name='Discret_Template')
discret_template.Convert()

#Analog template conversion
analog_template = Excel_to_Json(file_='DeTableV77.xlsx',
                                tab='Templates - Analogicos', 
                                column_to_drop= ['ElemName','Elem Type'],
                                output_file_name='Analog_Template')
analog_template.Convert()