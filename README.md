# MTE_Product_Search
Python Scripted for Web scrapping Product information from MTE Corp
Here is the steps and procedures in order to get the code working on your own personal computer

![image](https://user-images.githubusercontent.com/41027808/163736189-4b7c71e1-26c0-405a-aa00-82f8b61a3ff4.png)

# New Project Procedure
1. Create a python project using any IDE you like. This pythong script was creating using PyCharm Ide.
2. Create two new directories inside your project. Named "assets" and "chromedriver_win32" respectively. 
3. Download tci_logo_Csx_icon.ico and store it into the assets folder. 
4. Download Chrome and check your chrome version. To check your chrome vession simply go to the right hand corner and locate the three vertical symbol and click the symbols. Settings->About Chrome. 
5. Follow this link https://chromedriver.chromium.org/downloads and download chromedriver version based on your current Chrome Browser Version. 
6. Store the chromeDriver applications into chromedriver_win32 folder. 
7. Download sample excel file and store excel at your discretion.
8. In your interpter download the following Packages:
          
          * Seleniumum 4.10 
          * tkinterx 0.0.9
          * openpyxl 3.0.9
9. Verify the output of the GUI is as the image above:

Note: To represent or use your own personal logo simply insert an ico file into assets and change this line to represent your file name
         
         self.iconbitmap('./assets/tci_logo_Csx_icon.ico')

# GUI Procedure
This automation only scapes own product line at a time. Therefore verify that your input excel file only has one product line inside or else the program will not be successfull. 
1. Put your name and Date sinde the specific entry. Make sure to not use "\", "/", or "|" in the data. 
2. Click Insert file and insert your dignated file. (For testing use the sample excel file given).
3. Select webdriver from webdriver_chrom32 folder
4. Select new Excel locations in your computer
5. Slect Product Line in which your would like to scrape.
6. Open your input excel file and veriy that all products are as the same as the Selected Product Line, Verify that the sheet name is "Sheet1" and the product names are in Column 1 of "Sheet1"
7. Run Product Search by clicking Finished Settings

# Output
A new excel will be output in your selected excel locations. If no excel location was selected the file will be found inside your python project. This project will have the product name, descpriton, and Price of each product. 
