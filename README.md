#Assignment 4 

To run this tool run the duplex file
 

------------------------- Analysis of the use case-------------------------------

The goal is to calculate a price estimation based on the required ventilation loads
of a given building design and the cost of relevant ventilation products.  

 

The final BIM tool will automize the service of providing customers with a price estimation for both the air handling unit and ducts. 
The webpage will showcase the required inlet and exhaust per 
room to help the building designers estimate the ventilation loads.  

Furthermore, when the ventilation-duct-firm delivers the ducts, they are able to 
deliver in sorted packages – sorted/marked in terms of which space it is 
dimensioned by (space-wise). This will optimize the timeline of the construction 
workers because they can implement the system space by space, and therefore save 
time on figuring which ducts belongs to which space.  

The results show that BIM can facilitate communication of different information 
from an IFC-file to a user-friendly HTML-format.  

 

The purpose of the BIM tool is to assist the following stakeholders: 

- Building product manufacturer (selling ventilation ducts) 

- Architects, engineers, entrepreneurs, building owners (buying ventilation ducts) 

- Building constructors (installing the ventilation ducts) 

 

 

-----------------3B: Propose a design for a tool / workflow-------------------------- 
 

User workflow: 

The costumer uploads the BIM project as an IFC to the model folder.
 

Background processes for the tool: 

Then a IFC file is uploaded to the website, the python script will automatically 
run IFCOpenShell and extract the relevant data from the BIM model and output a list 
of all room types and the area of each room in the python console. 
The script links each room to the corresponding area, which makes it possible to calculate the total 
area of the building. The total area is used to determine the building type 
(residential or non-residential). 

The following calculations will link DS/EN13779 or 13142 requirements to the extracted 
data from the IFC file to produce a list of ventilations load per space. 
Assumptions about air flow velocity and duct shape is used to dimension of the ventilation 
ducts, whereafter an estimation of the total required length of ducts is made. 
Finally, a price offer is estimated based on the price of each duct size from a list in 
Excel (.exe) and the sizing of ventilation ducts.  

The results are converted to HTML format based on the python main file, which is depending 
on the CSS file which makes the layout and the Java script which enables the user 
interactions. 
The costumer is now presented with a table in HTML format including the estimated 
ventilation requirements, ducts size and price offer.  

 

---------------------------------3C: Information exchange------------------------------------

This BIM tool aims to assist designs at 200-300 Level Of Development (LOD) because the 
floorplan must be decided on, but the MEP design is not yet determined. 
The tool is therefore useful in the early design stages. 

We are interested in finding the room type and area of each room from the IFC file. 
From external sources we are using the DS/EN 13779 and DS/EN 13142 to determine 
requirements for the building.  

We are assuming the shape of the duct to be rectangular and an airflow to 2.5 m/s.  

We are furthermore assuming that all building above 400 m^2 are non-residential and 
all building below 400 m^2 are residential, since the ventilation requirements differs 
depending on whether the building is residential or not. We are aware that this 
assumption will lead to all apartment buildings being wrongly defined as non-residential 
due to the total floor area.  

The assumptions need more evidence to support further work.  

 

The use case manages the following data inputs: 

- The price of relevant products as an Excel sheet (from the product manufacturer) 

- DS/EN 13779 and 13142 requirements which are made as an Excel sheet

- A BIM model in IFC format 

- IFC entities extracted from the BIM model: Space.LongName, Area 

 

Outputs include an estimate of: 

- The required ventilation load in total and per room 

- The required ventilation duct size 

- The required length of ventilation ducts in total and per room 

- A price estimation 

 

The parameters impacting the output results are:  

- Building type (residential or non-residential) 

 

Requirements for the BIM tool to run: 

- A BIM model in IFC format with a floor plan including room names (In English) 

 

---------3D: Value - What is the potential improvement offered by this tool?-------------- 

**See attached picture (BusinessValue.svg) of the Use Case located in folder (img)

 

This BIM tool makes it easy and user-friendly for both product manufacturer and costumer 
to plan the installation of ventilation ducts for a given building design. 
It can be a possible tool for different MEP sector, so they avoid heavy calculations and 
get quick information of the ventilation and duct system in different buildings.  

The biggest business advantage is that the duct company can provide an extra service for 
its clients since they can get the price directly of the website on their specific 
project without any human interaction ;).  

The societal value is rooted in the OpenSource approach and attempts to assist small 
businesses selling ventilation ducts. 
A small company will not always have the resources to hire their own MEP engineers.  

The duct company does not need to purchase any software licenses, which can be very expensive 
too, so the economic value could be viewed as a societal value helping smaller companies 
reduce costs and stay in business.  

 

 

----------------------------------------3E: Delivery--------------------------------------------- 

The tool will optimally function as a web-based tool that can produce a list of dimensioned 
ducts and a total price estimation. The tool should mainly be used by ventilation ducts 
companies in the early stages of the design.  

The BIM-tool will be an add-on (or reference website) for the duct-company website. 
This will depend on our skills and security considerations running a local script through a HTML 
browser.  

The BIM tool can be edited by the firm, so the prices and duct sizes match the company-products. 
If we have the time, it could be cool to add the option of different types of ducts (not size, 
but material) – this should be doable.  

The customer/contractor will then upload a BIM-model (IFC) to the website (our BIM-tool) 
which will produce a price estimate. If the customer accepts the offer, the tool will optimally 
produce the specific duct-system design – we are not sure we are able to do this in time for 
deadline. Elsewise, we will use some rule of thumb when determining the amounts of ducts. 
This will be useful when presenting a price estimate when the contractor contacts the firm. 
Also, if the costumer decides to accept the offer, the ventilation-duct-firm, will be able to 
produce the different dimensioned ducts faster.  

The next step for us is to figure out if we can create the upload feature which will make the script 
run automatic with the rest of the script. Whether or not we can accomplish this, will determine 
the further workflow.  
