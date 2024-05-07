async function main(workbook: ExcelScript.Workbook): Promise<void> {
  // Check if the "Add POIs" worksheet exists.
  let sheet = workbook.getWorksheet("Get Custom Layers");
  if (!sheet) {
    console.log(`No worksheet named "Get Custom Layers" in this workbook.`);
    return;
  }

  const myInit = {
    method: "GET",
    headers: {
      "Authorization": "Key key=SCRUBBED",
      "Content-Type": "application/json"
    }
  };

  /**
  * An interface that matches the returned JSON structure.
  * The property names match exactly.
  */
  interface Layer {
    code: string;
    label: string;
    icon: string;
    sort: number;
    filter: FilterObject;
  }

  interface FilterObject {
    junction_type: string;
    predicates: PredicatesArray[];
  }

  interface PredicatesArray {
    operator: string;
    property: string;
    number_value: number;
    boolean_value: boolean;
    string_value: string;
  }

  try {
    const response = await fetch('https://platform.driveaxleapp.com/api/v1/poi_layers/custom', myInit);

    if (!response.ok) {
      throw new Error('Network response was not ok');
    }

    const layers: Layer[] = await response.json();

    // Create an array to hold the returned values.
    const rows: (string | boolean | number)[][] = [];

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

    // set column headers
    sheet.getRange('A1:K1').setValues([["Code", "Label", "Icon", "Sort","junction_type", "operator", "property", "number_value", "boolean_value", "string_value", "Active"]]);

    // Bold the column headers
    sheet.getRange('A1:K1').getFormat().getFont().setBold(true);

    // Add the data to the current worksheet, starting at "A2".
    const range = sheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
    range.setValues(rows);

  } catch (e) {
    console.log("Error: " + e);
  }
}
