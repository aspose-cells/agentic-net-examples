using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonExportDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Item");
            sheet.Cells["A2"].PutValue("Fruit");
            sheet.Cells["B2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Fruit");
            sheet.Cells["B3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Vegetable");
            sheet.Cells["B4"].PutValue("Carrot");

            // -------------------------------------------------
            // Add a data validation rule (list) to column B
            // -------------------------------------------------
            // Create a validation object and add it to the collection
            ValidationCollection validations = sheet.Validations;
            int validationIndex = validations.Add();                     // add a new validation and get its index
            Validation validation = validations[validationIndex];        // retrieve the validation object

            // Set validation type to List and provide the allowed values
            validation.Type = ValidationType.List;
            validation.InCellDropDown = true;                            // show drop‑down list
            validation.Formula1 = "Apple,Banana,Carrot,Tomato";          // comma‑separated list

            // Apply the validation to the desired area (B2:B4)
            CellArea area = new CellArea
            {
                StartRow = 1,   // zero‑based row index (row 2)
                EndRow = 3,     // row 4
                StartColumn = 1, // column B
                EndColumn = 1
            };
            validation.AddArea(area);

            // -------------------------------------------------
            // Configure JSON save options
            // -------------------------------------------------
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export the data as a JSON object (even if only one sheet)
                AlwaysExportAsJsonObject = true,
                // Include header row (first row) in the output
                HasHeaderRow = true,
                // Export empty cells as null (optional)
                ExportEmptyCells = true,
                // Do not use nested parent‑child structure for this simple table
                ExportNestedStructure = false,
                // Skip rows that are completely empty
                SkipEmptyRows = true
            };

            // -------------------------------------------------
            // Save the workbook as a JSON file
            // -------------------------------------------------
            string outputPath = "ExportedData.json";
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Workbook exported to JSON file: {outputPath}");
        }
    }
}