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
Uses the Eleos PUT POI Layers API to PUT Custom POI Layers into Eleos Platform Data.

Example:
```json
[
   {
       "sort": 0,
       "label": "Blue Beacon",
       "icon": "truckwash",
       "filter": {
           "predicates": [
               {
                   "string_value": "blue_beacon",
                   "property": "categories",
                   "operator": "EQUALS",
                   "number_value": null,
                   "boolean_value": null
               }
           ],
           "junction_type": "AND"
       },
       "code": "BLUE_BEACON"
   }
]
```

## 5. PUT Individual POIs
Uses the Eleos PUT POI API to send a set of POIs to the Eleos Platform Data.

Example:
```json
[
   {
       "name": "Blue Beacon",
       "location": {
           "longitude": "-84.48969",
           "latitude": "33.78546"
       },
       "full_address": "3105 Donald Lee Hollowell Pkwy NW Atlanta, GA 30318",
       "data": {
           "tags": [
               "cat"
           ],
           "street_name_suffix": "NW",
           "street_name_street_type": "Pkwy",
           "street_name_base_name": "Donald Lee Hollowell",
           "street_address": "3105 Donald Lee Hollowell NW Pkwy",
           "state": "Georgia",
           "routing_long": -84.48969,
           "routing_lat": 33.78546,
           "postal_code": "30318",
           "phone_number": "4047928996",
           "icon_ref": "cat",
           "house_number": "3105",
           "here_place_id": "840dn5bj-0983e55542cd43acb8fc4638ca2a206e",
           "here_location_id": "NT_RFVAkxDCcZm.aYFwprOOFD_zEDM1A",
           "fax_number": "404-799-7255",
           "display_long": -84.4886,
           "display_lat": 33.78664,
           "details": "Undercarraige Rinse",
           "country": "USA",
           "city": "Atlanta",
           "cat_scale": true
       },
       "categories": [
           "Car Wash-Detailing",
           "CAT Scale",
           "Truck Wash",
           "blue_beacon"
       ],
       "address": {
           "street_name": "Donald Lee Hollowell NW Pkwy",
           "state": "Georgia",
           "postal_code": "30318",
           "house_number": "3105",
           "country": "USA",
           "city": "Atlanta"
       }
   }
]
```


# Additional Information
Please remember to refer to the [Eleos POI API](https://dev.eleostech.com/platform/platform.html#tag/POIs) when making changes to the above scripts.
