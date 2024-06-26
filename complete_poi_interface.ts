async function main(workbook: ExcelScript.Workbook): Promise<void> {
    //Interfaces
  
    /************************************ Layer Interfaces **********************************************/
  
    // Interface that matches the returned JSON structure for a layer
    interface Layer {
      code: string;
      label: string;
      icon: string;
      sort: number;
      filter: FilterObject;
    }
  
    // Interface for the filter object with a layer
    interface FilterObject {
      junction_type: string;
      predicates: PredicatesArray[];
    }
  
    // Interface for the Predicates Array within the Filter Object
    interface PredicatesArray {
      operator: string;
      property: string;
      number_value: number | null;
      boolean_value: boolean | null;
      string_value: string | null;
    }
  
    /************************************** Individual POI Interfaces ****************************************/
  
    // Interface that matches the sent JSON for a POI
    interface poi {
      id: string;
      name: string;
      categories: string[];
      data: DataObject;
      full_address: string;
      location: LocationObject;
    }
  
    // Interface for the Data Object within a POI
    interface DataObject {
      details: string;
    }
  
    // Interface for the Location Object within a POI
    interface LocationObject {
      latitude: string;
      longitude: string;
    }

  
    /********************************************* Layer Functions ***********************************************/
    // Functions to create the spreadsheet

    // Function to convert row data to layer array
    function convertRowsToLayers(rowData: ExcelScript.Range): Layer {
      const numberValue = rowData.getCell(0, 7).getValue().toString();
      const booleanValue = rowData.getCell(0, 8).getValue().toString();
      const stringValue = rowData.getCell(0, 9).getValue().toString();
  
      const predicates: PredicatesArray[] = [];
  
      if (numberValue !== "") {
        predicates.push({
          operator: rowData.getCell(0, 5).getValue().toString(),
          property: rowData.getCell(0, 6).getValue().toString(),
          number_value: Number(numberValue),
          boolean_value: null, // Default value, as we are setting number_value
          string_value: null
        });
      } else if (booleanValue !== "") {
        predicates.push({
          operator: rowData.getCell(0, 5).getValue().toString(),
          property: rowData.getCell(0, 6).getValue().toString(),
          number_value: null,
          boolean_value: booleanValue.toLowerCase() === "true",
          string_value: null
        });
      } else if (stringValue !== "") {
        predicates.push({
          operator: rowData.getCell(0, 5).getValue().toString(),
          property: rowData.getCell(0, 6).getValue().toString(),
          number_value: null,
          boolean_value: null,
          string_value: stringValue
        });
      }
  
      const convertedLayer: Layer = {
        code: rowData.getCell(0, 0).getValue().toString(),
        label: rowData.getCell(0, 1).getValue().toString(),
        icon: rowData.getCell(0, 2).getValue().toString(),
        sort: Number(rowData.getCell(0, 3).getValue()),
        filter: {
          junction_type: rowData.getCell(0, 4).getValue().toString(),
          predicates: predicates
        }
      };
      return convertedLayer;
    }
  
    // fucntion to get active layers
    async function getLayers(sheet: ExcelScript.Worksheet) {
      try {
        const myInit = {
          method: "GET",
          headers: {
            "Authorization": "Key key=SCRUBBED",
            "Content-Type": "application/json"
          }
        };
  
        const response = await fetch('https://platform.driveaxleapp.com/api/v1/poi_layers/', myInit);
  
        if (!response.ok) {
          throw new Error('Network response was not ok');
        }
  
        const layers: Layer[] = await response.json();
  
        // Create an array to hold the returned values.
        const rows: (string | boolean | number | null)[][] = [];
  
        for (let each of layers) {
          // Determine if the row was returned, for example, if each.code exists
          const rowReturned = !!each.code;
  
  
          // Iterate over each predicate in the predicates array
          for (let predicate of each.filter.predicates) {
            // Push data along with the indicator into the rows array
            rows.push([
              each.code,
              each.label,
              each.icon,
              each.sort,
              each.filter.junction_type,
              predicate.operator,
              predicate.property,
              predicate.number_value,
              predicate.boolean_value,
              predicate.string_value,
              rowReturned
            ]);
          }
        }
  
  
        // Add the data to the current worksheet, starting at "A2".
        const range = sheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
        range.setValue(rows);
  
        // Set the content of the dropdown list.
        let validationCriteria: ExcelScript.ListDataValidation = {
          inCellDropDown: true,
          source: dropdownValues
        };
  
        let validationRule: ExcelScript.DataValidationRule = {
          list: validationCriteria
        };
        dataValidation.setRule(validationRule);
  
      } catch (e) {
        console.log("Error in getLayers: " + JSON.stringify(e));
      }
    }
  
    // Function to Upload Active Layers
    async function uploadLayers(sheet: ExcelScript.Worksheet) {
  
      const rows = sheet.getUsedRange().getRowCount(); // Get the last row with data
  
      // Collect data from each row
      for (let i = 1; i < rows; i++) {
  
        // see if the layer should be active
        const activeOrNot = sheet.getRange().getRow(i).getCell(0, 10).getValue();
  
        // if false delete the layer and continue
        if (activeOrNot === false) {
  
          // get layer code
          const code = sheet.getRange().getRow(i).getCell(0, 0).getValue().toString();
          console.log("code to delete: " + code);
  
          // delete the layer
          deletePOILayer(code);
  
          // clear the range (the row)
          sheet.getRange().getRow(i).delete(ExcelScript.DeleteShiftDirection.up);
  
          continue;
        }
        else {
          const row = sheet.getRange().getRow(i);
  
          const layers = convertRowsToLayers(row);
  
          const myInit = {
            method: "PUT",
            headers: {
              "Authorization": "Key key=SCRUBBED",
              "Content-Type": "application/json"
            },
            body: JSON.stringify(layers)
          };
  
          try {
            const response = await fetch('https://platform.driveaxleapp.com/api/v1/poi_layers/', myInit);
  
            if (!response.ok) {
              let errorMessage = "Error: " + response.status;
              if (response.body !== null) {
                errorMessage += " " + await response.text();
              }
              console.log("error in upload Layers")
              throw new Error(errorMessage);
            } else {
              console.log("it worked!");
            }
  
          } catch (e) {
            console.log("Error in upload layers #2: " + e);
          }
        }
      }
    }
  
    // function to delete POI Layers
    async function deletePOILayer(code: string) {
      const myInit = {
        method: "DELETE",
        headers: {
          "Authorization": "Key key=SCRUBBED",
          "Content-Type": "application/json"
        }
      };
  
      const response = await fetch('https://platform.driveaxleapp.com/api/v1/poi_layers//' + code, myInit);
  
      if (!response.ok) {
        let errorMessage = "Error in delete layer: " + response.status;
        if (response.body !== null) {
          errorMessage += " " + await response.text();
        }
        console.log("error in delete layer #2");
        throw new Error(errorMessage);
      } else {
        console.log("it worked! " + code + " deleted");
      }
    }
  
    /************************************** Individual POI Functions ***************************************/
  
    // Function to convert row data to poi object
    function convertRowToPoi(rowData: ExcelScript.Range): poi {
      try {
        // Extract data from the row and create poi object
        let poiObject: poi = {
          id: rowData.getCell(0, 0).getValue().toString(),
          name: rowData.getCell(0, 1).getValue().toString(),
          categories: rowData.getCell(0, 2).getValue().toString().split(',').map(category => category.trim()),
          full_address: rowData.getCell(0, 3).getValue().toString(),
          // Initialize data and location objects
          data: {
            details: rowData.getCell(0, 4).getValue().toString()
          },
          location: {
            latitude: rowData.getCell(0, 5).getValue().toString(),
            longitude: rowData.getCell(0, 6).getValue().toString(),
          }
        };
        return poiObject;
      } catch (error) {
        console.log("Error in convertRowToPoi:", error);
        throw error; // Re-throw the error to propagate it further
      }
    }
  
  
    // function to upload POIs
    async function uploadPOIs(poiSpreadsheet: ExcelScript.Worksheet) {
      try {
  
        const usedRange = poiSpreadsheet.getRange().getUsedRange();
        const rows = usedRange.getRowCount();
  
        console.log("usedRange: " + usedRange);
        console.log("rows: " + rows);
  
        for (let i = 1; i < rows; i++) {
          const poiRow = usedRange.getRow(i); // Get the range for the current row
  
          const poiObject = convertRowToPoi(poiRow); // Convert row data to poi object
  
          const myInit = {
            method: "PUT",
            headers: {
              "Authorization": "Key key=SCRUBBED",
              "Content-Type": "application/json"
            },
            body: JSON.stringify([poiObject]) // Convert poiArray to JSON string
          };
  
          console.log(JSON.stringify([poiObject])); // Log the poi array
  
          const response = await fetch('https://platform.driveaxleapp.com/api/v1/pois', myInit);
  
          if (!response.ok) {
            let errorMessage = "Error: " + response.status;
            if (response.body !== null) {
              errorMessage += " " + await response.text();
            }
            throw new Error(errorMessage);
          } else {
            console.log("it worked!");
          }
        }
      } catch (error) {
        console.log("Error in uploadPOIs: " + error);
      }
    }
  

    /********************************** Spreadsheet Setup ********************************/
  
    // Check if the "Add POIs" worksheet exists.
    const sheet = workbook.getWorksheet("Custom Layer Interface");
    if (!sheet) {
      console.log(`No worksheet named "Custom Layer Interface" in this workbook.`);
      return;
    }
  
    // Check if the "Add POIs" worksheet exists.
    const poiSpreadsheet = sheet.getRange("M2").getValue().toString();
    const poisheet = workbook.getWorksheet(poiSpreadsheet);
    if (!poisheet) {
      console.log(`No worksheet named` + poiSpreadsheet + `in this workbook.`);
      return;
    }
  
    let cellN2Value = sheet.getRange("N2").getValue();
  

    /********************************** Spreadsheet Formatting ********************************/
  
    const columnRange = sheet.getRange('A1:K1');
    // set column headers
    columnRange.setValues([["Code", "Label", "Icon", "Sort", "junction_type", "operator", "property", "number_value", "boolean_value", "string_value", "Active"]]);
  
    // Bold the column headers
    columnRange.getFormat().getFont().setBold(true);
  
    // set update headers
    sheet.getRange('M1:O1').setValues([["POI Spreadsheet", "Action", "Time Updated"]]);
  
    // Bold the update headers
    sheet.getRange('M1:O1').getFormat().getFont().setBold(true);
  
    // Define the range where you want to apply the in-cell dropdown
    let dropdownRange = sheet.getRange("N2");
    const dataValidation = dropdownRange.getDataValidation();
  
    // Set the values for the dropdown list
    let dropdownValues = "Update,Refresh";

    /********************************** Main ********************************/
  
    // If the sheet is blank or cell N2 is set to "Refresh"
    if (sheet.getRange("A1").getValue() === "" || cellN2Value === "Refresh") {
      getLayers(sheet);
    }
    else if (cellN2Value === "Update") {
      // upload edited layers
      await uploadLayers(sheet);
    }
  
    // time of update
    const timeReturned = new Date(Date.now()).toISOString();
    sheet.getRange('O2').setValue(timeReturned);
  
    // format cells
    sheet.getUsedRange().getFormat().autofitColumns();
  
    // Call the function to upload POIs
    await uploadPOIs(poisheet);
  
  
  }
  
