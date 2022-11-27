#written by group 18
import pandas as pd
import ifcopenshell
import os.path
import time
import numpy as np


##############################################################################
#
##############################################################################

model = ifcopenshell.open('model\duplex.ifc')
#model = ifcopenshell.open('model\Officebuilding_v3.ifc')

    
#Now we pull out the class type (here the spaces) which we are interested in
spaces = model.by_type('IfcSpace')


#following loop generates two arrays: (1)space types, (2)areas per space. 
space_area = []
space_type = []   
for space in spaces:
    for definition in space.IsDefinedBy:
        if definition.is_a('IfcRelDefinesByProperties'):
            property_set = definition.RelatingPropertyDefinition

       
#finding type-of-name byt IFC file.
#TYPE 1 made from Revit                           
        if property_set.Name == "PSet_Revit_Dimensions":            
            for property in property_set.HasProperties:     
                if property.Name == "Area":
                    #print(property.NominalValue, space.LongName)
                    ifc_area = str(property.NominalValue)
                    area_per_space = float(((ifc_area.split("("))[1])[:-1])
                    space_area.append(area_per_space)
                    space_type.append(space.LongName)
                    #print(space.LongName,":", round(area_per_space,1), "m2")
        
        
       
#finding type-of-name byt IFC file.
#TYPE 2 made from ArciCAD
        if property_set.Name == "BaseQuantities": 
            for property in property_set.Quantities:
                #print(property.get_info())
                if property.Name == "GrossFloorArea":
                    area_per_space = float(property.AreaValue)
                    space_area.append(area_per_space)
                    space_type.append(space.Name)
#                    print(space.Name,":", round(ifc_area,1), "m2")     

##############################################################################
#                              TOTAL AREA
##############################################################################

#finds total amount of area 
total_area = sum(space_area)


#delete if we call other type of xax
xax = np.stack((space_type, space_area), axis=1)

#print(xax)

##############################################################################
#                               SMALL BUIDLINGS
##############################################################################

if total_area <= 600:        
    
    #following loop generates array, which describes the amount of exhaust per space
    space_exhaust = []
    for i in range(len(xax)):    
        if 'Bathroom' in xax[i,0]:
            space_exhaust.append(float(15))
        elif 'Kitchen' in xax[i,0]:
            space_exhaust.append(float(20))
        else:
              space_exhaust.append(float(0))   
    
    #calculates total exhaust (DELETE LATER WHEN USE IN OTHER)
    total_exhaust = sum(space_exhaust) 
    
    #following loop generates array, which tells if space is a Bedroom or a living room.
    # If true, it returns the area of the space. 
    space_findRooms = []
    for i in range(len(xax)):
        if 'Bedroom' in xax[i,0]:
            space_findRooms.append(float(xax[i,1]))
        elif 'Living Room' in xax[i,0]:
            space_findRooms.append(float(xax[i,1]))
        else:
            space_findRooms.append(0)
    
    total_areaSupply = np.sum(space_findRooms)
    #print(type(total_areaSupply))     
    
    #loop finding Supply of spaces
    space_Supply = []
    for i in range(len(xax)):
        calc_Supply = (space_findRooms[i]/total_areaSupply)*(total_exhaust)
        space_Supply.append(calc_Supply)


    
    ###PRICE###
    price_ex = pd.read_excel('.\input\Ventilation_prices.xlsx')
    
    duct_price = []
    for w in range(len(space_exhaust)):
        #only one type of ducts 
        duct_price.append(round((space_exhaust[w] + space_Supply[w]) * 0.3 * price_ex.Price_per_meter[1],1))
        
    #####
   
    header_info = ["Type", "Area", "Exhaust", "Supply", "Duct Price"]
    all_info = np.stack((space_type, space_area, space_exhaust, space_Supply, duct_price), axis=1)   

##############################################################################
#                         NON-RESIDENTIAL BUIDLINGS
##############################################################################
else:         
    req = pd.read_excel('.\input\Ventilation_req.xlsx')
    
#    req = pd.read_excel('.\input\Ventilation_req.xlsx', sheet_name='USE_CASE')

    
    space_exhaust_1 = []
    space_exhaust_2 = []
    space_exhaust_3 = []
    space_exhaust_4 = []
    
    total_other_1 = []
    total_other_2 = []
    total_other_3 = []
    total_other_4 = []
    
    total_other = [total_other_1, total_other_2, total_other_3, total_other_4]
    all_arrays = [space_exhaust_1, space_exhaust_2, space_exhaust_3, space_exhaust_4]
    for i in range(len(xax)):    
        if 'Single Office' in xax[i,0]:       
            a = [req.a[0], req.a[1], req.a[2], req.a[3]] #l/s per m2
            b = [req.b[0], req.b[1], req.b[2], req.b[3]] #l/2 per Person
            for j in range(4):
                q = ((float(xax[i,1]) / float(10.0)) * float(b[j])) + (float(xax[i,1]) * float(a[j]))
                exhaust = all_arrays[j]
                exhaust.append(round(q,1))
    
        elif 'Landscape office' in xax[i,0]:       
            a = [req.a[4], req.a[5], req.a[6], req.a[7]] #l/s per m2
            b = [req.b[4], req.b[5], req.b[6], req.b[7]] #l/2 per Person
            for j in range(4):
                q = ((float(xax[i,1]) / float(15.0)) * float(b[j])) + (float(xax[i,1]) * float(a[j]))
                exhaust = all_arrays[j]
                exhaust.append(round(q,1))
    
        elif 'Confrence' in xax[i,0]:       
            a = [req.a[8], req.a[9], req.a[10], req.a[11]] #l/s per m2
            b = [req.b[8], req.b[9], req.b[10], req.b[11]] #l/2 per Person
            for j in range(4):
                q = ((float(xax[i,1]) / float(2.0)) * float(b[j])) + (float(xax[i,1]) * float(a[j]))
                exhaust = all_arrays[j]
                exhaust.append(round(q,1))
    
        elif 'Audiotorium' in xax[i,0]:       
            a = [req.a[12], req.a[13], req.a[14], req.a[15]] #l/s per m2
            b = [req.b[12], req.b[13], req.b[14], req.b[15]] #l/2 per Person
            for j in range(4):
                q = ((float(xax[i,1]) / float(0.75)) * float(b[j])) + (float(xax[i,1]) * float(a[j]))
                exhaust = all_arrays[j]
                exhaust.append(round(q,1))
    
        elif 'Resturant' in xax[i,0]:       
            a = [req.a[16], req.a[17], req.a[18], req.a[19]] #l/s per m2
            b = [req.b[16], req.b[17], req.b[18], req.b[19]] #l/2 per Person
            for j in range(4):
                q = ((float(xax[i,1]) / float(1.5)) * float(b[j])) + (float(xax[i,1]) * float(a[j]))
                exhaust = all_arrays[j]
                exhaust.append(round(q,1))
    
        elif 'Class' in xax[i,0]:       
            a = [req.a[20], req.a[21], req.a[22], req.a[23]] #l/s per m2
            b = [req.b[20], req.b[21], req.b[22], req.b[23]] #l/2 per Person
            for j in range(4):
                q = ((float(xax[i,1]) / float(2.0)) * float(b[j])) + (float(xax[i,1]) * float(a[j]))
                exhaust = all_arrays[j]
                exhaust.append(round(q,1))
            
        else:
            a = [2,2,2,2] #l/s per m2
            b = [2,2,2,2] #l/2 per Person
            for j in range(4):
                q = ((float(xax[i,1]) / float(1.0)) * float(b[j])) + (float(xax[i,1]) * float(a[j]))
                exhaust = all_arrays[j]
                totalother = total_other[j]
                exhaust.append(round(q,1))
                totalother.append(round(q,1))

#calc supply
    space_supply_1 = []
    space_supply_2 = []
    space_supply_3 = []
    space_supply_4 = []
    all_arrays_s = [space_supply_1, space_supply_2, space_supply_3, space_supply_4]
   
    for i in range(len(xax)):
        if 'Single Office' in xax[i,0]:       
            for j in range(4):
                q = all_arrays[j][i] + (sum(total_other[j]) / (len(all_arrays[0])-len(total_other[0])))
                supply = all_arrays_s[j]
                supply.append(round(q,1))    

        elif 'Landscape office' in xax[i,0]:       
            for j in range(4):
                q = all_arrays[j][i] + (sum(total_other[j]) / (len(all_arrays[0])-len(total_other[0])))
                supply = all_arrays_s[j]
                supply.append(round(q,1))   

        elif 'Confrence' in xax[i,0]:       
            for j in range(4):
                q = all_arrays[j][i] + (sum(total_other[j]) / (len(all_arrays[0])-len(total_other[0])))
                supply = all_arrays_s[j]
                supply.append(round(q,1))   

        elif 'Audiotorium'  in xax[i,0]:       
            for j in range(4):
                q = all_arrays[j][i] + (sum(total_other[j]) / (len(all_arrays[0])-len(total_other[0])))
                supply = all_arrays_s[j]
                supply.append(round(q,1))   

        elif  'Resturant' in xax[i,0]:       
            for j in range(4):
                q = all_arrays[j][i] + (sum(total_other[j]) / (len(all_arrays[0])-len(total_other[0])))
                supply = all_arrays_s[j]
                supply.append(round(q,1))   

        elif  'Class' in xax[i,0]:       
            for j in range(4):
                q = all_arrays[j][i] + (sum(total_other[j]) / (len(all_arrays[0])-len(total_other[0])))
                supply = all_arrays_s[j]
                supply.append(round(q,1))   
                
        else:
            for j in range(4):
                supply = all_arrays_s[j]
                supply.append(0)
                
    
    ###PRICE###
    price_ex = pd.read_excel('.\input\Ventilation_prices.xlsx')
   
    duct_price_1 = []
    duct_price_2 = []
    duct_price_3 = []
    duct_price_4 = []

    duct_price_large = [duct_price_1, duct_price_2, duct_price_3, duct_price_4]
    for m in range(len(space_supply_1)):
        
        for z in range(4):            
            (duct_price_large[z]).append(round((all_arrays[z][m] + all_arrays_s[z][m]) * 0.3 * price_ex.Price_per_meter[1],1))
          
    
    # header_info = ["space_type", "space_area" ,"Category 1: supply [l/s]" ,"space_exhaust_1","Category 2: supply [l/s]" ,"space_exhaust_2", "Category 1: supply [l/s]", "space_exhaust_3", "s4", "space_exhaust_4"]
    # all_info = np.stack((space_type, space_area, space_supply_1, space_exhaust_1, space_supply_2, space_exhaust_2, space_supply_3, space_exhaust_3, space_supply_4, space_exhaust_4), axis=1)  

    header_info = ["space_type", "space_area" ,"Category 1: supply [l/s]" ,"Category 1: exhaust [l/s]", "Duct price [kr]","Category 2: supply [l/s]" ,"Category 2: exhaust [l/s]", "Duct price [kr]", "Category 3: supply [l/s]" ,"Category 3: exhaust [l/s]", "Duct price [kr]", "Category 4: supply [l/s]" ,"Category 4: exhaust [l/s]", "Duct price [kr]"]
    all_info = np.stack((space_type, space_area, space_supply_1, space_exhaust_1, duct_price_1, space_supply_2, space_exhaust_2, duct_price_2, space_supply_3, space_exhaust_3, duct_price_3, space_supply_4, space_exhaust_4, duct_price_4), axis=1)


##############################################################################
#                               AIR
##############################################################################


    total_exhaust = [sum(all_arrays[0]), sum(all_arrays[1]), sum(all_arrays[2]), sum(all_arrays[3])]
    
    #The list puts out: 1. the total air change in m3/s, 2. Height, 3. Width, 4. Lenght, 5. Volume, 6.(???Måske prisen her???)
    c1 = []
    c2 = []
    c3 = []
    c4 = []
    classes1_4 = [c1, c2, c3, c4]
    #for i in range(4):
    
    for i in range(4):    
        list_of_type=[]
        a = round(total_exhaust[i]/1000, 2)
          
        if a <= 0.7:
            values = (a,0.97,0.97,2.16,round(0.97*0.97*2.16,2))
            (classes1_4[i]).append(values)
            
        elif a > 0.7 and a <= 1:
            values = (a,1.12,1.12,2.16,round(1.12*1.12*2.16,2))
            (classes1_4[i]).append(values)
            
        elif a > 1 and a <= 1.5:
           values = (a,1.27,1.27,2.46,round(1.27*1.27*2.46,2))
           (classes1_4[i]).append(values)
        
        elif a > 1.5 and a <= 2:
            values = (a,1.42,1.42,2.46,round(1.42*1.42*2.46,2))
            (classes1_4[i]).append(values)
           
        elif a > 2 and a <= 2.5:
            values = (a,1.57,1.57,2.76,round(1.57*1.57*2.76,2))
            (classes1_4[i]).append(values)
    
        elif a > 2.5 and a <= 3:
            values = (a,1.78,1.87,3.06,round(1.78*1.87*3.06,2))
            (classes1_4[i]).append(values)
    
        elif a > 3 and a <= 4:
            values = (a,2.02,2.02,2.91,round(2.02*2.02*2.91,2))
            (classes1_4[i]).append(values)
    
        elif a > 4 and a <= 5:
            values = (a,2.24,2.17,3.28,round(2.24*2.24*3.28,2))
            (classes1_4[i]).append(values)
    
        elif a > 5 and a <= 6:
            values = (a,2.54,2.17,3.21,round(2.54*2.17*3.21,2))
            (classes1_4[i]).append(values)
    
        elif a > 6 and a <= 7.5:
            values = (a,2.84,2.37,3.96,round(2.84*2.37*3.96,2))
            (classes1_4[i]).append(values)
    
        elif a > 7.5 and a <= 9:
            values = (a,3.14,2.59,4.26,round(3.14*2.59*4.26,2))
            (classes1_4[i]).append(values)
    
        elif a > 9 and a <= 11:
            values = (a,3.44,2.89,4.56,round(3.44*2.89*4.56,2))
            (classes1_4[i]).append(values)
    
        elif a > 11 and a <= 15:
            values = (a,4.34,3.19,5.01,round(4.34*3.19*5.01,2))
            (classes1_4[i]).append(values)
    
        elif a > 15:
            values = (a,4.94,3.49,5.3,round(4.94*3.49*5.31,2))
            (classes1_4[i]).append(values)
    
    header_info_ahu = ["info","r1", "r2", "r3", "r4"]
    row_info_ahu = ["m2/s","L","B","H","Volumen"]
    
    all_info_ahu = np.stack((row_info_ahu,(np.array(classes1_4[0]))[0],(np.array(classes1_4[1]))[0],(np.array(classes1_4[2]))[0],(np.array(classes1_4[3]))[0]), axis=1)  


##############################################################################
#                                   EXCEL
##############################################################################


## convert your array into a dataframe
df = pd.DataFrame(all_info, columns = header_info)
#df_ahu= pd.DataFrame(all_info_ahu)
#df_ahu= pd.DataFrame(all_info_ahu, columns = header_info_ahu)


## save to xlsx file in the input folder 

#filepath = r'C:\Users\Maria\OneDrive - Danmarks Tekniske Universitet\Skrivebord\BIM\Ventilation_Results.xlsx'
filepath = 'input\Ventilation_Results.xlsx'

#filepath_ahu = r'C:\Users\Maria\OneDrive - Danmarks Tekniske Universitet\Skrivebord\BIM\ahu.xlsx'

#filepath = r'C:\Users\Maria\OneDrive - Danmarks Tekniske Universitet\Skrivebord\my_excel_file.xlsx'

#"C:\Users\Maria\OneDrive - Danmarks Tekniske Universitet\Skrivebord\BIM"

df.to_excel(filepath, index=False)
#df_ahu.to_excel(filepath_ahu, index=False)



##############################################################################
#
##############################################################################




def modelLoader(name):

    ''' 
        load the IFC file 
    '''
        
    model_url = "model/duplex.ifc"
    start_time = time.time()

    if (os.path.exists(model_url)):
        model = ifcopenshell.open(model_url)
        print("\n\tFile    : {}.ifc".format(name))
        print("\tLoad    : {:.2f}s".format(float(time.time() - start_time)))
        
        start_time = time.time()
        writeHTML(model,name)
        print("\tConvert : {:.4f}s".format(float(time.time() - start_time)))
        
    else:
        print("\nERROR: please check your model folder : " +model_url+" does not exist")

def writeHTML(model,name):

    ''' 
        write the HTML entities 
    '''
    
    # parent directory - put in setting file?
    parent_dir = "output/"
    #create an HTML file to write to
    
    if (os.path.exists("output/"+name))==False:
        path = os.path.join(parent_dir, name)
        os.mkdir(path)
    
    f_loc="output/"+name+"/index.html"
    f = open(f_loc, "w")
    cont=""
    
    # ---- START OF STANDARD HTML
    cont+=0*"\t"+"<html>\n"
    # ---- ADD HEAD
    cont+=1*"\t"+"<head>\n"
    # ---- ADD HTMLBUILD CSS - COULD ADD OTHERS HERE :)
    cont+=2*"\t"+"<link rel='stylesheet' href='../css/html-build.css'>"
    
    # ---- ADD BACKGROUND IMAGE
    #cont+=2*"\t"+'<style type="text/css">\n  body { \n background-image: url(../../../../../../../Downloads/Atlas_forrest.jpg); \n}">'
    #cont+=3*"\t"+ '</style>' 
    
    # ---- END LINK
    cont+=2*"\t"+'</link>'
    
    
    # ---- ADD JS
    cont+=2*"\t"+'<script src="../js/html-build.js"></script>'
    cont+=2*"\t"+'<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.6.0/jquery.js"></script>'
    
    # ---- TEXT STYLE
    cont+=2*"\t"+'<meta charset="utf-8">'
    
    # ---- ADD TITLE
    cont+=1*"\t"+'<title>VENTILATION</title>'
    
    # ---- CLOSE HEAD
    cont+=1*"\t"+"</head>\n"
    
    
    # ---- ADD BODY
    cont+=1*"\t"+'<body onload=\"main()\">\n'  
    

    
    
    # ---- ADD Table
    
    
    # ---- ADD CUSTOM HTML FOR THE BUILDING HERE
    cont+=writeCustomHTML(model)
    
    
    # ---- ADD
    # ---- CLOSE BODY AND HTML ENTITIES
    cont+=0*"\t"+"</body>\n"   
    cont+=0*"\t"+"</html>\n"

    # ---- WRITE IT OUT
    f.write(cont)
    f.close()

    # ---- TELL EVERYONE ABOUT IT
    print("\tSave    : "+f_loc)

def writeCustomHTML(model):

    ''' 
        write the custom HTML entities 
    '''
    
    custom=""

    # ---- DEFINE THE TABEL   
    custom+=2*"\t"+"<view->\n"
    custom+=3*"\t"+"<center>\n"
    #custom+=4*"\t"+'<table width="383" height="276" border="0,5">'
    custom+=5*"\t"+'<caption>'
    custom+=6*"\t"+'<h1>Estimated Ventilation Loads</h1> <h2> Inlet and outlet per room</h2>'


        # ----- READ FROM EXCEL FILE IN FOLDER 'MODEL* 
    data = pd.read_excel('.\input\Ventilation_Results.xlsx') 
    data = data.round(decimals=1)
    custom+=data.to_html()

    #custom+=4*"\t"+'<\table>'
    #custom+=3*"\t"+'<\center>' 
    #custom+=1*"\t"+'<\view->' 
    #custom+=2*"\t"+'<center>' 
    #custom+=2*"\t"+'<img src="../../img/VENTI.JPG" alt=""/>' 
    #custom+=2*"\t"+'<\center>' 
    

    return custom

# def writeCustomHTML(model):

#     ''' 
#         write the custom HTML entities 
#     '''
    
#     custom=""

#     # ---- DEFINE THE TABEL   
#     custom+=2*"\t"+"<view->\n"
#     custom+=3*"\t"+"<center>\n"
#     #custom+=4*"\t"+'<table width="383" height="276" border="0,5">'
#     custom+=5*"\t"+'<caption>'
#     custom+=6*"\t"+'<h1>Estimated Ventilation Loads</h1> <h2> Inlet and outlet per room</h2>'
#     custom+=5*"\t"+'</caption>'

#         # ----- READ FROM EXCEL FILE IN FOLDER 'MODEL* 
#     data = pd.read_excel('.\input\Ventilation_Results.xlsx') 
#     data = data.round(decimals=1)
#     custom+=data.to_html()

#     #custom+=4*"\t"+'<\table>'
#     #custom+=3*"\t"+'<\center>' 
#     #custom+=1*"\t"+'<\view->' 
#     #custom+=2*"\t"+'<center>' 
#     #custom+=2*"\t"+'<img src="../../img/VENTI.JPG" alt=""/>' 
#     #custom+=2*"\t"+'<\center>' 
    

#     return custom