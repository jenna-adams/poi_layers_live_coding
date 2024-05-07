async function main(workbook: ExcelScript.Workbook): Promise<void> {

  // Check if the "Add POIs" worksheet exists.
  let sheet = workbook.getWorksheet("Add POIs");
  if (!sheet) {
    console.log(`No worksheet named "Add POIs" in this workbook.`);
    return;
  }

  /**
   * An interface that matches the JSON structure.
   * The property names match exactly.
   */
  interface poi {
    id: string;
    name: string;
    categories: string[];
    data: DataObject;
    full_address: string;
    location: LocationObject;
  }

  interface DataObject {
    details: string;
  }

  interface LocationObject {
    latitude: string;
    longitude: string;
  }

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

  async function uploadPOIs() {
    try {

      const usedRange = sheet.getRange().getUsedRange();
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
        console.log("Error in uploadPOIs: "+ error);
      }
    }
  


  // Call the function to upload POIs
  await uploadPOIs();
}
