// Title: Create an Excel workbook, add list‑type data validation to a column, and export the sheet as indented JSON using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a new Workbook, fills it with sample rows, applies a list‑type validation to column A, and saves the first worksheet as a formatted JSON file with headers using Aspose.Cells. | Demonstrate how to set up JsonSaveOptions to export a single‑sheet workbook as a JSON object, include empty cells as null, and indent the output for readability.
// Common Searches: asp.net add list validation to Excel column with Aspose.Cells | how to export an Aspose.Cells worksheet to pretty printed JSON in C# | Aspose.Cells JsonSaveOptions export empty cells as null example | C# save workbook with data validation as JSON file using Aspose.Cells | configure header row handling when converting Excel to JSON with Aspose.Cells
// Tags: list validation Aspose.Cells | export worksheet to JSON Aspose.Cells | JsonSaveOptions indentation | header row handling JSON Aspose.Cells | apply data validation before JSON export

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonExportDemo
{
    // Shows how to instantiate a Workbook, populate sample data, add a list‑type validation to column A, configure JsonSaveOptions (header row, object output, empty cells as null, indentation), and save the sheet as a formatted JSON file.
    class Program
    {
        static void Main(string[] args)
        {
            // 1. Instantiate a new workbook
            Workbook workbook = new Workbook();

            // 2. Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // 3. Populate some sample data (including a header row)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Fruits");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Vegetables");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["A4"].PutValue("Beverages");
            sheet.Cells["B4"].PutValue(45);

            // 4. Add a data validation rule (list validation) to column A (Category)
            //    The allowed values are: Fruits, Vegetables, Beverages
            ValidationCollection validations = sheet.Validations;
            int validationIndex = validations.Add();                     // create a new validation
            Validation validation = validations[validationIndex];        // retrieve the validation object

            // Set validation type to List and provide the comma‑separated list of allowed values
            validation.Type = ValidationType.List;
            validation.InCellDropDown = true;                           // show drop‑down arrow
            validation.Formula1 = "Fruits,Vegetables,Beverages";

            // Apply the validation to the range A2:A4 (rows 2‑4 in column A)
            CellArea area = new CellArea
            {
                StartRow = 1,      // zero‑based index, row 2
                EndRow = 3,        // row 4
                StartColumn = 0,   // column A
                EndColumn = 0
            };
            validation.AddArea(area);

            // 5. Configure JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export the data as a JSON object (since we have a single sheet)
                AlwaysExportAsJsonObject = true,
                // Treat the first row as header names
                HasHeaderRow = true,
                // Export empty cells as null (optional)
                ExportEmptyCells = true,
                // Indent the output for readability
                Indent = "  "
            };

            // 6. Save the workbook as a JSON file
            string outputPath = "WorkbookWithValidation.json";
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Workbook saved to JSON file: {outputPath}");
        }
    }
}
