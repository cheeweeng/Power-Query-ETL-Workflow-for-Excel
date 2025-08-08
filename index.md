This project will use Power Query to set up connections to various data sources, pre-process the data, and load the finalized data into the Excel for further analysis and dashboarding.  
<div style="display: flex; gap: 10px;">
  <img width="400" height="600" alt="Image1" src="https://github.com/user-attachments/assets/a02a8a6e-d6fb-48d5-bace-21a617373af6" />
  <img width="500" height="471" alt="Image2" src="https://github.com/user-attachments/assets/1b4ec1fb-7ab0-4fa3-9fd7-d0cf51cf2d45" />
</div> 

Open a new workbook and get data from the data tab to import the customer dataset. In the Navigator pane, click on 'Edit' (Excel 2016) or 'Transform' (Excel 2019 above) to open up the Power Query Editor.  

### Change data type  
<img width="645" height="340" alt="Image" src="https://github.com/user-attachments/assets/d8dcb43d-ffef-4518-b83a-b3719c021d18" />  

Change the data type of customer_id from numeric to text.    

<img width="555" height="440" alt="Image" src="https://github.com/user-attachments/assets/e6010dd4-87c9-428b-9d75-8aa778b22381" />    

To replace a value in a column, from the Transform tab click on replace values.  A dialogue box will pop up, enter the 'Value to Find' and 'Replace With'.  Do the same for Male.  

Now, lets add a conditional column.
<img width="1000" height="500" alt="Image" src="https://github.com/user-attachments/assets/1985a736-4761-4205-9183-e09bfd5eeb35" />  

Click on the column to apply the condition, in this case 'education'. In the Add column tab, click on 'Conditional Columns'. 
A dialogue box will pop up, give the new column a name 'education_level' and enter the If, ELIF and ELSE conditions as shown.  

<img width="700" height="305" alt="Image" src="https://github.com/user-attachments/assets/903f6f11-b621-4c54-a667-9076698cd740" />    

Next, lets add a age column by clicking on the Date icon in Add column tab and selecting Age.

<img width="1090" height="318" alt="Image" src="https://github.com/user-attachments/assets/dd3cf4b3-4406-4c9d-9152-07f2345b6713" />  
A new column called Age will be created but we have to transform the column to the proper data type.  From the Trasnform tab, click on duration and select Total Years.
Still highlighting the Age column, change the data type to whole number.  

### Merging columns  
<img width="800" height="458" alt="Image" src="https://github.com/user-attachments/assets/a2b58a3e-8950-4195-b4e9-0ab11f711675" />  

The customers' address are seperated into various columns, we will merge the relevant details into a single column. 
Select the required columns to be merged, then in Add column tab, click on Merge Columns.  From the dialogue box, enter the desired seperator and named the new column.

### Close and Load to Workbook
<img width="363" height="500" alt="Image" src="https://github.com/user-attachments/assets/b65dd88a-56ea-49ba-92cf-7590a2cbbdd7" />  

From the Home tab, select Close and Load to Workbook, then followed by Only Create Connection.  
By creating connection only, the query does not load any data into Excel table or Power BI model.  It creates a connection object that can be used later.

### Get data from folder
<div style="overflow: auto;">
  <img src="https://github.com/user-attachments/assets/92fcdddf-473a-4727-a054-d84dadf3fdda" width="378" height="418" style="float: left; margin-right: 15px;" />

  <div style="text-align: right;">
    If the data to be analyze are organized by weekly, monthly or yearly, then these files need to be combined before the analysis can be performed.  
    For this project, the transaction data are in yearly format and saved in the transaction folder that contains two files for year 2021 and 2022.  
    Important to note that the files must contain the same number of columns with the same name and the sequence of the columns must be the same.
  </div>
</div>


<img width="355" height="390" alt="Image" src="https://github.com/user-attachments/assets/9fcbb31d-f9f8-41e5-9f93-62450da588db" />  
<p>Similar to earlier step, change the customer_id data type to text and then remove the source.name column.  </p>


<img width="700" height="380" alt="Image" src="https://github.com/user-attachments/assets/24da33fb-f3df-4e39-9a6f-fcb098548e48" />   

Let's add a new column to find the margin of each product by subtracting the product cost from the product retail price.  To subtract columns, the sequence of selecting the columns will affect the outcome. In this example, the product retail price column must be selected first, followed by the product cost column, in order to get the margin of each product.  

<img width="500" height="320" alt="Image" src="https://github.com/user-attachments/assets/15674b05-4d5c-43ec-913c-c599a07def22" />    

Next, let's add a new column using Multiply. This time, we will find the amount spent by multiplying the product retail price by the quantity (sold). For Multiply, the sequence in selecting the columns do not matter.  


<div style="display: flex; gap: 10px;">
<img width="600" height="350" alt="Image" src="https://github.com/user-attachments/assets/3cb6964e-fcf7-4983-8530-65d73303448a" />
<img width="380" height="419" alt="Image" src="https://github.com/user-attachments/assets/995d3313-aff0-4e54-ab4c-1e7d873ea39c" />
</div>  
To add a Quarter column, select the transaction date field and from the add column tab, click the Date icon.  From the dropdown, select Quarter and then Quarter of Year.  
Next, transform the Quarter column by adding a prefix 'Q'.  

Now the pre-processing work is completed, close and load to workbook and only create connection.  

<div style="display: flex; gap: 10px;">
<img width="300" height="425" alt="Image" src="https://github.com/user-attachments/assets/110b25e3-9581-4efc-8849-f23229487d39" />
<img width="550" height="515" alt="Image" src="https://github.com/user-attachments/assets/af9f7e09-517b-4108-8e2a-57a54e11c241" />
</div>
Combine the two queries transactions and customers by selecting the matching column, which is customer_id. Note the data type of this column must be the same for both connection, otherwise it will return an error.  

<img width="500" height="470" alt="Image" src="https://github.com/user-attachments/assets/b58969b6-46dd-403d-b759-b7f12ce42b1c" />  
<p>From the Power Query Editor, a new column customers will be created. Click on the top right corner of the column and select
the desired columns you want to add to the transaction table. Also give the merged connection a name. </p> 

<img width="470" height="320" alt="Image" src="https://github.com/user-attachments/assets/6d4ebb7f-f0e7-497b-8d06-592557b04926" />  
<p>Reduce dimensionality of the dataset by remove unnecessary columns, in this case, the various address columns.  </p>

<img width="330" height="600" alt="Image" src="https://github.com/user-attachments/assets/a0de4bd3-3bde-4680-9fd2-c312b0c8dad2" />
<p>To remove duplicates, click on the icon to the left of the first column and select remove duplicates from the drop down menu.</p>

<img width="500" height="630" alt="Image" src="https://github.com/user-attachments/assets/93a4c1cb-b4f4-433b-8fb4-e73efa2209fa" />
<p>Now the data pre-processing work has been done on the combined queries, it is time to close and load as a Table</p>

<img width="900" height="327" alt="Image" src="https://github.com/user-attachments/assets/2709e66c-68f2-4350-9588-f52b73dd3725" />
<p>A Table will be created and ready for further analysis and visualization work</p>


