// Title: C# – Create an Excel workbook with list validation and export it to JSON using Aspose.Cells
// Description: This example shows how to build a new Workbook, fill cells A1:B4 with sample data, apply a list‑based validation to column B (B2‑B10) with a drop‑down list, configure JsonSaveOptions (nested hierarchy, skip empty rows, header row), and save the worksheet as output.json.
// Keywords: Aspose.Cells C# export to JSON | list validation Excel Aspose | JsonSaveOptions nested JSON | skip empty rows JSON export | Excel header row to JSON | .NET workbook validation example | Aspose.Cells data validation list
// Common Searches: Aspose.Cells add list validation C# | export Excel sheet to JSON with Aspose.Cells | JsonSaveOptions ExportNestedStructure example | how to skip empty rows when saving JSON from Excel | C# code to create workbook and save as JSON
// Developer Intent: Create a workbook, enforce a dropdown list on a column, and generate a JSON file with hierarchical structure and clean data.
// Use Cases: Transform an Excel template with validated dropdowns into a JSON payload for a web API. | Produce a compact JSON representation of a worksheet for front‑end consumption, preserving column headers and removing blank rows. | Generate parent‑child JSON data from Excel while ensuring input values conform to a predefined list.
// AI Prompts: Generate C# code that adds a list validation to column C and exports the sheet to JSON with ExportNestedStructure disabled. | Show how to read the output.json file produced by Aspose.Cells and deserialize it into a List<T> in .NET. | Explain how to modify JsonSaveOptions to include cell formulas and comments in the exported JSON.

using System;
using Aspose.Cells;

// This example shows how to build a new Workbook, fill cells A1:B4 with sample data, apply a list‑based validation to column B (B2‑B10) with a drop‑down list, configure JsonSaveOptions (nested hierarchy, skip empty rows, header row), and save the worksheet as output.json.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data with a header row
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Item");
        sheet.Cells["A2"].PutValue("Fruit");
        sheet.Cells["B2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Fruit");
        sheet.Cells["B3"].PutValue("Banana");
        sheet.Cells["A4"].PutValue("Vegetable");
        sheet.Cells["B4"].PutValue("Carrot");

        // Add a list‑type data validation to column B (Item)
        Validation validation = sheet.Validations[sheet.Validations.Add()];
        validation.Type = ValidationType.List;          // List validation
        validation.InCellDropDown = true;               // Show drop‑down arrow
        validation.Formula1 = "Apple,Banana,Carrot,Tomato"; // Allowed values

        // Apply the validation to cells B2:B10
        CellArea area = new CellArea
        {
            StartRow = 1,    // Row index is zero‑based (B2)
            StartColumn = 1, // Column B
            EndRow = 9,      // B10
            EndColumn = 1
        };
        validation.AddArea(area);

        // Configure JSON save options
        JsonSaveOptions saveOptions = new JsonSaveOptions
        {
            ExportNestedStructure = true, // Export as parent‑child hierarchy
            SkipEmptyRows = true,         // Omit empty rows
            HasHeaderRow = true           // First row contains headers
        };

        // Save the workbook as a JSON file using the configured options
        string outputPath = "output.json";
        workbook.Save(outputPath, saveOptions);
    }
}
