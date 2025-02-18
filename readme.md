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

Example Body:
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

Example Body:
```json
[
    {
        "id": "112",
        "name": "Blue Beacon #112",
        "location": {
            "longitude": "-84.339417",
            "latitude": "33.659323"
        },
        "full_address": "4170 Old McDonough Rd, Conley, GA 30288",
        "data": {
            "details": "Phone Number: 217-342-4303"
        },
        "categories": [
            "blue_beacon"
        ]
    }
]
```


# Additional Information

## Juction Type and Predicates
These fields can cause some confusion, here's some background.
The `junctionType` and `predicates` fields are used to define how filtering rules are applied to determine whether a POI (Point of Interest) should be included in the layer.

Here's a breakdown:
- **`junctionType`**: This defines how multiple predicates should be combined. It has two possible values:
  - `"AND"`: All predicates must be true for the POI to be included.
  - `"OR"`: At least one predicate must be true for the POI to be included.

- **`predicates`**: This is an array of filter conditions that determine which POIs belong to the layer. Each predicate typically has:
  - A `fieldName`: The attribute being checked (e.g., `"category"`, `"name"`, etc.).
  - An `operator`: The comparison type (e.g., `"EQUALS"`, `"CONTAINS"`, `"GREATER_THAN"`).
  - A `value`: The expected value for the condition.

### Example  
Let's say you want to create a POI layer that includes truck stops **AND** have more than 50 parking spaces. You’d use:

```json
{
  "junctionType": "AND",
  "predicates": [
    {
      "fieldName": "category",
      "operator": "EQUALS",
      "value": "Truck Stop"
    },
    {
      "fieldName": "parkingSpaces",
      "operator": "GREATER_THAN",
      "value": 50
    }
  ]
}
```

Alternatively, if you want to include all **Truck Stops OR Rest Areas**, you’d use:

```json
{
  "junctionType": "OR",
  "predicates": [
    {
      "fieldName": "category",
      "operator": "EQUALS",
      "value": "Truck Stop"
    },
    {
      "fieldName": "category",
      "operator": "EQUALS",
      "value": "Rest Area"
    }
  ]
}
```



Please remember to refer to the [Eleos POI API](https://dev.eleostech.com/platform/platform.html#tag/POIs) when making changes to the above scripts.
