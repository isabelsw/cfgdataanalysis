<a name="readme-top"></a>

<br />
<div align="center">
  <p><img src="https://i.giphy.com/coxQHKASG60HrHtvkt.webp" width="480"></p>
<br />
<h3 align="center">BASIC DATA ANALYSIS</h3>

<!-- PROJECT LINKS -->
  <p align="center">
    <br />
    <a href="https://github.com/isabelsw/cfgdataanalysis"><strong>Go Straight to the Docs »</strong></a>
    <br />
    <br />
    <a href="https://github.com/isabelsw/cfgdataanalysis/issues/new?labels=bug&template=bug-report---.md">Report a bug</a>
    ·
    <a href="https://github.com/isabelsw/cfgdataanalysis/issues/new?labels=enhancement&template=feature-request---.md">Request a feature</a>
  </p>
</div>

<p>&nbsp;</p>

<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li><a href="#about-the-project">About The Project</a></li>
    <li><a href="#built-with">Built With</a></li>
    <li><a href="#features">Features</a></li>
    <li><a href="#setup">Setup</a></li>
    <li><a href="#roadmap">Roadmap</a></li>
    <li><a href="#contributing">Contributing</a></li>
    <li><a href="#licence">Licence</a></li>
    <li><a href="#contact">Contact</a></li>
    <li><a href="#acknowledgements">Acknowledgements</a></li>
    <li><a href="#copyright-notices">Copyright Notices</a></li>
  </ol>
</details>

<p>&nbsp;</p>

<!-- ABOUT THE PROJECT -->
## <a name="about-the-project"></a>About The Project

This is a solution to the assignment "Spreadsheet Analysis" as outlined in the course "Introduction to Python & Apps" by <a href="https://codefirstgirls.com/">Code First Girls Ltd</a>. It entails basic analysis of fictional sales data using both Excel and Python. See <a href="#features">Features</a> for an overview of the methods and libraries/modules used.
<p>&nbsp;</p>
<p align="right">(<a href="#readme-top">back to top</a>)</p>



## <a name="built-with"></a>Built With

| Python | PyCharm | pandas | Matplotlib | Seaborn | Microsoft Excel 365 |
|----------|----------|----------|----------|----------|----------|
| <img src="https://github.com/devicons/devicon/blob/master/icons/python/python-original.svg" title="Python"  alt="Python" width="55" height="55"/> | <img src="https://github.com/devicons/devicon/blob/master/icons/pycharm/pycharm-original.svg" title="pycharm" alt="pycharm" width="55" height="55"/>| <img src="https://github.com/devicons/devicon/blob/master/icons/pandas/pandas-original.svg" title="Pandas" alt="Pandas" width="55" height="55"/>| <img src="https://github.com/devicons/devicon/blob/master/icons/matplotlib/matplotlib-original.svg" title="mpl" alt="mpl" width="55" height="55"/> | <img src="https://seaborn.pydata.org/_images/logo-mark-lightbg.svg" title="Seaborn" alt="Seaborn" width="55" height="55"/>| <img src="https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white"/> |

<p>&nbsp;</p>

## <a id=“features”>Features</a>

* Use different libraries/modules for simple analysis of data from an Excel workbook: 
  - pandas
  - openpyxl
  - seaborn
  - matplotlib
* Find the total sum of values and the minimum and maximum values
* Use pct_change for working out periodic changes as a percentage
* Create a simple lineplot for showing periodic changes as a percentage
* Import an Excel file and then write results to a new Excel file
* Assign data to different spreadsheets upon writing to the new Excel file
<p>&nbsp;</p>
<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- SETUP -->
## <a name="setup"></a>Setup

### Prerequisites

* <a href="https://www.python.org/downloads/">Install Python<a/> and/or <a href="https://www.jetbrains.com/pycharm/download/?section=windows#section=windows">Install PyCharm<a/> by Jet Brains (you can opt for the free and open-source **community edition**). Make sure you are downloading a Python 3 version.
* Get a subscription for Microsoft 365 and download the latest verstion of <a href="https://www.microsoft.com/en-gb/microsoft-365/excel">MS Excel</a>.
* See <a href="https://realpython.com/pycharm-guide/">Pycharm Guide</a> for creating a new Py project.
* Download the file sales.csv via the <a href="https://github.com/isabelsw/cfgdataanalysis">repository's main page</a>.


* PIP package

1) See if PIP is already installed on your system. Write in either Windows OS shell or macOS terminal:

   ### Install
     ```sh
     pip help
     ```
  
     ```sh
     pip3 help
     ```
  
     ```sh
     python3 -m pip help
     ```
  
3) If it is not installed, write in shell or terminal:
  
     ```sh
     python3 get-pip.py
     ```

  N.B. Make sure you have added pip.exe to the Environment Variables (Path).
  
* pandas library; openpyxl module; seaborn library; matplotlib
  
  There are 2 ways, EITHER: 
  
  - Write in PyCharm IDE install [library name]. E.g.,

        install pandas
  
   - OR Write in shell or terminal pip install [library name]. For macOS terminal, if that instruction does not work, you can write pip3 install [library name]. E.g.,

        ```sh
        pip install pandas
        ```
<p>&nbsp;</p>
<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- ROADMAP -->
## <a name="roadmap"></a>Roadmap 

If you want to have a go yourself, I have provided the **guide** below. You will create **3 separate Py files** that reflect each stage of the process. 
- N.B. From this point on, **Python** will be abbreviated as **Py**.

PY FILE 1

* Create a new **.py** file in your project folder in PyCharm (if you have not already done so). Name the file.
* Import **openpyxl and math**.
  - N.B. It is good practice to import your libraries/modules at the top; however, I have imported pandas as pd later in the code so that it is shown within its immediate context.
* Import **sales.csv** file (<a href="https://github.com/isabelsw/cfgdataanalysis">sales.csv</a>); otherwise, create your own spreadsheet with similar quantitative data.

<p>&nbsp;</p>

![PyCharm_table_sales_csv](https://github.com/isabelsw/cfgdataanalysis/assets/170036120/a3423aa4-092d-4f71-83fd-64f12b2c9c52)

<p>&nbsp;</p>

* **Create a function** named read_data using **def** to retrieve the sales and months from the spreadsheet.
* Use the **def method** to create a function that makes a Py dictionary from the data for monthly sales (from sales.csv). 
* Collect **all the sales** from each month into a single list using the **list() function**.
* Get the **sales total** (across all months) using the **sum() function**.
* Find the **largest and smallest values** in the sales data using the **min() and max() functions**.
* Find the **months with the lowest/highest number** of sales by using the **dictionary get() method** and writing the **keys as max_value and min_value**.
* Find the **average** value of sales using the **lens() function** (returns the number of values for sales). Then continue the formula by writing **total / number_of_sales** (total divided by the number of values returned by the lens function).
* Using **math.ceil() method** (from the math module), **round the number** of average sales that was calculated in the previous step. You do not have to do this, but it ensures better readability of the table that we will be creating later.
* **Import pandas as pd** to read sales.csv and create a dataframe. Use **pct_change() method** to retrieve sales data and calculate the monthly change in sales from one month to the next as a percentage. Multiply **pct_change by 100** (* 100) to complete the formula. NaN will appear in the table for Jan if you do not use the **fillna() function**; this is because there is no value for the previous month (in the sales data) with which to calculate the change in sales for Jan. Fill none with the value 0 to make it clear that there is no monthly-change data for this month.
* Create a **pandas dataframe** for the output. Then use **ExcelWriter** and **df.to_excel** to write the results (**df**) to a new workbook and spreadsheet. Write the parameter **index=False** to remove the surplus column.
* This step is **optional**: you can load the existing workbook and **change its formatting** so that the size of the cells fit the values better. I have used **from openpyxl import load_workbook** to load the workbook and **from openpyxl.utils import get_column_letter** for this task.
* **Save and close** the workbook.
* **Print** the results with the **string format() method** so that they appear in the console (you will need these values for writing to the new Excel sheet; see below). **/n** in the print function adds a line break. Then write **run()** to complete def run().

<p>&nbsp;</p>

![PyCharm_Console_Output_basicdataanalysis](https://github.com/isabelsw/cfgdataanalysis/assets/170036120/c9bf5604-2665-41ea-b7d2-1185cf8ecf72)

<p>&nbsp;</p>

PY FILE 2

* Import **pandas as pd**.
* Read the Excel file (created via the previous steps) into a **pandas dataframe** object with the **pd.read_excel function**. 
* With your results from PY FILE 1, add the new data as a **dictionary of lists** and create a **dataframe** for it with the **pd.DataFrame function**.
* Write the **dataframe** to a new Excel spreadsheet using **pd.Excelwriter**. Name your sheet in the parameters and write **index=false** to remove the generated extra column.

PY FILE 3

* Import: **seaborn as sns**; **pandas as pd**; **matplotlib as plt**.
* Specify columns from which to retrieve the data. Read this data into a pandas dataframe via the **pd.read_excel function**. 
* **Print** this data.

<p>&nbsp;</p>

![PyCharm_Console_Output_graph](https://github.com/isabelsw/cfgdataanalysis/assets/170036120/c1df8b1f-f280-4173-9684-e783943ec1cd)

<p>&nbsp;</p>

![Figure_1_cfgspreadsheetanalysis](https://github.com/isabelsw/cfgdataanalysis/assets/170036120/58ed36fb-f665-466a-b0ef-9b71ad74e4b0)

<p>&nbsp;</p>

* Create a simple lineplot with the **sns.lineplot() function**.
* Show this lineplot by writing **plt.show()**. You can then save this lineplot to a local folder.


See the [open issues](https://github.com/isabelsw/cfgdataanalysis/issues) for a full list of proposed features and known issues.
<p>&nbsp;</p>
<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- CONTRIBUTING -->
## <a name="contributing"></a>Contributing

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement".
Don't forget to give the project a star! Thanks!

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`) 
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request
<p>&nbsp;</p>
<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- LICENCE -->
## <a name="licence"></a>Licence

Distributed under the MIT Licence. See "LICENSE.txt" for more information.
<p>&nbsp;</p>
<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- CONTACT -->
## <a name="contact"></a>Contact

My profile - <a href="https://github.com/isabelsw">isabelsw</a>

Project link - [https://github.com/isabelsw/cfgdataanalysis](https://github.com/isabelsw/cfgdataanalysis)
<p>&nbsp;</p>
<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- ACKNOWLEDGMENTS -->
## <a name="acknowledgements"></a>Acknowledgements

* The project brief was created by <a href="https://codefirstgirls.com/">Code First Girls Ltd</a>. A special thanks goes to my instructors Vanny and Andrew on the "Introduction to Python & Apps" course.
* <a href="https://github.com/othneildrew/Best-README-Template/blob/master/BLANK_README.md?plain=1">README template</a> was created by <a href="https://github.com/othneildrew">Othneil Drew</a>.
* GIF courtesy of <a href="https://giphy.com/gifs/coxQHKASG60HrHtvkt">GIPHY</a>.
<p>&nbsp;</p>
<p align="right">(<a href="#readme-top">back to top</a>)</p>


<!-- COPYRIGHT -->
## <a name="copyright-notices"></a>Copyright Notices

<!-- Please note that I have used a JPEG file for showing the individual logos of the languages used in this project for stability. I have provided the links to the companies below. MS Office does not let other parties use its logo "in any manner" :( -->

* <a href="https://www.jetbrains.com/">PyCharm</a> logo: Copyright © 2000-2024 JetBrains s.r.o. JetBrains and the JetBrains logo are registered trademarks of JetBrains s.r.o.
* <a href="https://www.python.org/">Python</a> logo: Copyright © 2001-2024 Python Software Foundation.
* <a href="https://pandas.pydata.org/">pandas</a> logo: Copyright © 2008 AQR Capital Management, LLC, Lambda Foundry, Inc. and PyData Development Team.
* <a href="https://seaborn.pydata.org/index.html">Seaport</a> logo: Copyright © Matthias Bussonnier and Seaport.
* <a href="https://matplotlib.org/stable/">matplotlib</a> logo: Copyright ©  2002–2012 John Hunter, Darren Dale, Eric Firing, Michael Droettboom and the Matplotlib development team; 2012–2024 The Matplotlib development team.
<p>&nbsp;</p>
<p align="right">(<a href="#readme-top">back to top</a>)</p>
