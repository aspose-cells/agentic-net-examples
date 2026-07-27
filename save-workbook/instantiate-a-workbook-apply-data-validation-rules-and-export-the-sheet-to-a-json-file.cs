using System;
using Aspose.Cells;

namespace AsposeCellsDataValidationToJson
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data (optional, helps visualize the JSON output)
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("Item1");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("Item2");
            worksheet.Cells["B3"].PutValue(20);

            // -------------------------------------------------
            // Add a data validation rule to cell A2 (list type)
            // -------------------------------------------------
            // Get the validations collection of the worksheet
            ValidationCollection validations = worksheet.Validations;

            // Add a new validation and obtain its index
            int validationIndex = validations.Add();

            // Retrieve the validation object
            Validation validation = validations[validationIndex];

            // Set validation type to List and define the allowed values
            validation.Type = ValidationType.List;
            validation.Formula1 = "Option1,Option2,Option3";

            // Apply the validation to cell A2 (row 1, column 0)
            CellArea area = new CellArea
            {
                StartRow = 1,    // zero‑based index, row 2 in Excel
                StartColumn = 0, // column A
                EndRow = 1,
                EndColumn = 0
            };
            validation.AddArea(area);

            // -------------------------------------------------
            // Configure JSON save options
            // -------------------------------------------------
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export as a nested (parent‑child) JSON structure
                ExportNestedStructure = true,
                // Skip rows that contain no data
                SkipEmptyRows = true,
                // Include header row in the JSON output
                HasHeaderRow = true
            };

            // Save the workbook as a JSON file using the configured options
            string outputPath = "ValidatedData.json";
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Workbook saved to JSON file: {outputPath}");
        }
    }
}