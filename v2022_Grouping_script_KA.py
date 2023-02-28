import iesve   # the VE api
import pandas as pd
import numpy as np

#Get Model data from IES
project = iesve.VEProject.get_current_project() #get project from IES
models = project.models[0] #get models from IES
bodies = models.get_bodies(False) #get bodies from IES

excel_path = 'revB_IES_automatization_interface_KA.xlsm' #Get user selection from Excel saved in the same folder
user_selection = pd.read_excel(excel_path,sheet_name = "Main") #read homepage in excel
selected_typology = user_selection.iat[0,0] #read user defined selection in excel from drop down list. Don't forget to save excel file after changing the input

#Create room groups in IES based on user selection in Excel and read the keywords
rg = iesve.RoomGroups() #create room groups 
scheme_index = rg.create_grouping_scheme(str(selected_typology)) #create group scheme index
df = pd.read_excel(excel_path,sheet_name = str(selected_typology))

#get a list of all keys and values from user input. First column should contain Group Names and the next column should contain the search terms. 
HL_group_names = df.iloc[:,0].tolist()
HL_search_terms = df.iloc[:,1].tolist()  
print(HL_group_names)
print(HL_search_terms)
 
#Iterate through all the room, search their names against search terms and assign them to matching group names if a match is found.
for id, name in enumerate(HL_search_terms):
    group_list=[]  
    colour = tuple(np.random.choice(range(256), size=3))
    ies_group = rg.create_room_group(scheme_index, HL_group_names[id],colour)  
    
    
    for body in bodies:
        room_data = body.get_room_data(0)
        # Print all the room data   
        data = room_data.get_general()
        # check all names to categorize 
        if "+" in str(name):
            a,b = str(name).split("+")
            if a.lower().strip() in (data['name']).lower() and b.lower().strip() in (data['name']).lower():
                group_list.append(body)
        else:            
            for word in str(name).split(","):
                if word.lower().strip() in (data['name']).lower():
                    group_list.append(body)
    rg.assign_rooms_to_group(scheme_index,ies_group, group_list)
