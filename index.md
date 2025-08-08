This project will use Power Query to set up connections to various data sources, pre-process the data, and load the finalized data into the Excel for further analysis and dashboarding.  
<div style="display: flex; gap: 10px;">
  <img width="430" height="700" alt="Image1" src="https://github.com/user-attachments/assets/a02a8a6e-d6fb-48d5-bace-21a617373af6" />
  <img width="500" height="471" alt="Image2" src="https://github.com/user-attachments/assets/1b4ec1fb-7ab0-4fa3-9fd7-d0cf51cf2d45" />
</div> 

Open a new workbook and get data from the data tab to import the customer dataset. In the Navigator pane, click on 'Edit' (Excel 2016) or 'Transform' (Excel 2019 above) to open up the Power Query Editor.  

### Change data type  
<img width="745" height="361" alt="Image" src="https://github.com/user-attachments/assets/d8dcb43d-ffef-4518-b83a-b3719c021d18" />  

Change the data type of customer_id from numeric to text.    

<img width="595" height="472" alt="Image" src="https://github.com/user-attachments/assets/e6010dd4-87c9-428b-9d75-8aa778b22381" />    

To replace a value in a column, from the Transform tab click on replace values.  A dialogue box will pop up, enter the 'Value to Find' and 'Replace With'.  Do the same for Male.  

Now, lets add a conditional column.
<img width="1177" height="556" alt="Image" src="https://github.com/user-attachments/assets/1985a736-4761-4205-9183-e09bfd5eeb35" />  

Click on the column to apply the condition, in this case 'education'. In the Add column tab, click on 'Conditional Columns'. 
A dialogue box will pop up, give the new column a name 'education_level' and enter the If, ELIF and ELSE conditions as shown.  

<img width="797" height="305" alt="Image" src="https://github.com/user-attachments/assets/903f6f11-b621-4c54-a667-9076698cd740" />    

Next, lets add a age column by clicking on the Date icon in Add column tab and selecting Age.

<img width="1198" height="318" alt="Image" src="https://github.com/user-attachments/assets/dd3cf4b3-4406-4c9d-9152-07f2345b6713" />  
A new column called Age will be created but we have to transform the column to the proper data type.  From the Trasnform tab, click on duration and select Total Years.
Still highlighting the Age column, change the data type to whole number.  

### Merging columns  
<img width="900" height="508" alt="Image" src="https://github.com/user-attachments/assets/a2b58a3e-8950-4195-b4e9-0ab11f711675" />  

The customers' address are seperated into various columns, we will merge the relevant details into a single column. 
Select the required columns to be merged, then in Add column tab, click on Merge Columns.  From the dialogue box, enter the desired seperator and named the new column.

### Close and Load to Workbook
<img width="363" height="508" alt="Image" src="https://github.com/user-attachments/assets/b65dd88a-56ea-49ba-92cf-7590a2cbbdd7" />  

From the Home tab, select Close and Load to Workbook, then followed by Only Create Connection.  
By creating connection only, the query does not load any data into Excel table or Power BI model.  It creates a connection object that can be used later.

### Get data from folder
<img width="378" height="418" alt="Image" src="https://github.com/user-attachments/assets/92fcdddf-473a-4727-a054-d84dadf3fdda" />  






