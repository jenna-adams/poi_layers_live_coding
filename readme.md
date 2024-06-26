# Excel Script Toolbox for POI Control

Welcome to the Excel Script repository designed to facilitate control over Eleos Platform Data, specifically POI Layers, and Individual POIs, using Excel.


# Overview

The scripts provided in this repository leverage Excel Script, which is essentially TypeScript tailored for Excel automation. With these scripts, users can streamline their workflow, manage Eleos Platform Data, manipulate POI Layers, and handle Individual POIs directly within Excel.

## Getting Started

To start using the scripts in this repository, follow these steps:

1.  Download the repository to your local machine.
2.  Ensure you have the necessary permissions and access to run scripts within Excel.
3.  Open Excel and navigate to the Script Editor.
4.  Copy and paste the desired script code from this repository into the Script Editor. Editing where necessary.
5.  Save and execute the script.

For more detailed instructions on how to use Excel Script, refer to the official [Excel Script documentation](https://learn.microsoft.com/en-us/office/dev/scripts/).

# Scripts

## 1. Complete POI interface 
This script was demoed live during the Eleos Developer Summit of 2024. This script will pull all current custom layers, and upload/delete all specified POI layers, then upload all individual POIs.
It requires the starter **xlsx** file.

## 2. Starter **xlsx** File
Starter Excel file needed to run the complete POI interface file.

## 3. GET Custom POI Layers
Uses the Eleos GET POI Layers API, to display that information in a blank spreadsheet.

## 4. PUT Custom POI Layers
Uses the Eleos PUT POI Layers API to PUT Custom POI Layers into Eleos Platform Data

## 5. PUT Individual POIs
Uses the Eleos PUT POI API to send a set of POIs to the Eleos Platform Data.


# Additional Information
Please remember to refer to the [Eleos POI API](https://dev.eleostech.com/platform/platform.html#tag/POIs) when making changes to the above scripts.

