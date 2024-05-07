async function main(workbook: ExcelScript.Workbook): Promise<void> {
  // get the current sheet
  let sheet = workbook.getActiveWorksheet();

  /**
   * An interface that matches the body JSON structure.
   * The property names match exactly.
   */
  interface LayerArray {
    code: string;
    label: string;
    icon: string;
    sort: number;
    filter: FilterArray;
  }

  interface FilterArray {
    junction_type: string;
    predicates: PredicateObject[];
  }

  interface PredicateObject {
    operator: string;
    property: string;
    string_value: string;
  }

  // Function to convert row data to layer array
  function convertRowsToLayers(rowData: ExcelScript.Range): LayerArray {

    const layer: LayerArray = {
      code: rowData.getCell(0, 0).getValue().toString(),
      label: rowData.getCell(0, 1).getValue().toString(),
      icon: rowData.getCell(0, 2).getValue().toString(),
      sort: Number(rowData.getCell(0, 3).getValue()),
      filter: {
        junction_type: rowData.getCell(0, 4).getValue().toString(),
        predicates: [{
          operator: rowData.getCell(0, 5).getValue().toString(),
          property: rowData.getCell(0, 6).getValue().toString(),
          string_value: rowData.getCell(0, 7).getValue().toString()
        }]
      }
    };
    return layer;
  }

  async function uploadLayers() {
    const rows = sheet.getUsedRange().getRowCount(); // Get the last row with data

    // Collect data from each row
    for (let i = 1; i < rows; i++) {
      const row = sheet.getRange().getRow(i);
    

    const layers = convertRowsToLayers(row);
    console.log("layer: ", layers)

    const myInit = {
      method: "PUT",
      headers: {
        "Authorization": "Key key=SCRUBBED",
        "Content-Type": "application/json"
      },
      body: JSON.stringify(layers)
    };

    try {
      const response = await fetch('https://platform.driveaxleapp.com/api/v1/poi_layers/custom', myInit);

      if (!response.ok) {
        let errorMessage = "Error: " + response.status;
        if (response.body !== null) {
          errorMessage += " " + await response.text();
        }
        throw new Error(errorMessage);
      } else {
        console.log("it worked!");
      }

    } catch (e) {
      console.log("Error: " + e);
    }
  }
  }

  // Call the function to upload layers
  await uploadLayers();
}
