# DAEN 690: Wine Production Dashboard for Pearmund Cellars

As part of the Data Analytics Engineering Program at George Mason University, the DAEN 690 capstone course provides students with the opportunity to solve real world problems using the knowledge and skills gained from their course work. Our team, team Sommelier Analytics, was tasked with helping Pearmund Cellars and it's sister wineries forecast the quantity of wine to be produced next season. Our partners, Chris Pearmund and John Memomli, also wanted to better understand their data to help explain and quantify not only their business but the situtation surround the forcasted production.

To accomplish this a Wine Production Dashobard was created with Excel. This dashboard contains three things:
* A database to hold historical data and a process to input new data.
* Numerous business intelligence tools to draw key insights from the data and visualize them through tables and graphs.
* A pipeline to predict demand for their wines as well as to adjust the values based on the existing business situtation and inventory levels to forecast production.

#### Database

The database was a native function to Excel and required no extra implementation. Initial datasets were provided by the partners which were cleaned and transformed to meet the needs of the project and placed in seperate sheets on the dashboard. This allowed the dashboard to host all relevant and required data and allowed the other features to easily reference information. User input forms were also created so new data can be inputted.

To ensure privacy all proprietary data has been stripped from the dashboard. Examples of valid data forms are shown in each dataset.

#### Business Intelligence

Numerous Business Intelligence tools were created using VBA scripts and button controls, both native to Excel. The various scripts utilize user selections from drop-down lists to make proper selections on which dataset to use. These scripts also perform the necessary calculations and out the results in either a table or graph for easy visualization.

The VBA scripts for each BI tool are included seperatly from the dashboard in the "VBA Scripts" folder.

#### Demand & Production Pipeline

The Demand & Production Pipeline using historical case movement data fed into an ARIMA model to predict demand. This demand value is then adjusted based on existing inventory levels and a desired months left of inventory range.

To implement an ARIMA model, the VBA script calls the Python executable to run a Python script containing the model. This code queries the correct data, validates it, runs the model and then outputs results back to Excel. This process is done for a choice list of wines selected by the partners, generating 12 months of predictions. Python was needed because Excel currently does not have to ability to run ARIMA models at the speed and ease needed for this project.

To adjust the predicted demand to forecast production values the total amount of predicted demand is added the current inventory level and a months left of inventory value is calculated based on average monthly case movement. If the months left value falls outside a user specifed range, the value is adjsuted so it falls within this range. This value is then the amount of wine in cases the wineries should produce in the coming season.

This code is detailed both in the "VBA Scripts" folder as well as the "Python" folder.

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes.

### Prerequisites

Note that this project was designed to run on Windows OS due to the feature limitations seen in MacOS versions of Excel.

You will need the following software to run this project on your machine:

[Excel](https://www.microsoft.com/en-us/microsoft-365/excel) - The framework of the project, any version after 2016 will suffice

[Python](https://www.python.org/) - Used to run the ARIMA model to predict demand, any version after 3.9 will suffice

Within Python the following packages need to be installed:
- pandas
- numpy
- datetime
- tqdm
- sklearn
- scipy
- dateutil
- openpyxl
- pmdarima
- sys

### Installing

To get the project running on a local machine use the following steps:

1) Download and install all software detailed in the Prerequisites section
2) Download the "Wine Production Dashboard" folder which contains all files needed
3) Open the WineProductionDashboard.xlsm and change the necessary path references relative to the local machine
4) Repeat step 3 for the ARIMAbackend.py file

## Authors

* **Nebojsa Hrnjez** - *Product Owner*
* **Bhargav Sai Barala** - *Scrum Master*
* **Rohan Jonnakuti** - *Developer*
* **Kshetragna Rao Gummadapu** - *Developer*

## License

The data from Pearmund Cellars was licences exclusively for the use within the project and therefore is not uploaded to GitHub

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details
