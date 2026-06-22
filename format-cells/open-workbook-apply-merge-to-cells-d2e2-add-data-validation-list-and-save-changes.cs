using System;
using Aspose.Cells;

namespace AsposeCellsMergeAndValidation
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path if needed)
            Workbook workbook = new Workbook(); // creates a new workbook; you can also use new Workbook("input.xlsx") to load

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Merge cells D2:E2 (zero‑based indices: row 1, column 3)
            // -------------------------------------------------
            // firstRow = 1 (row 2), firstColumn = 3 (column D), totalRows = 1, totalColumns = 2
            worksheet.Cells.Merge(1, 3, 1, 2);

            // -------------------------------------------------
            // Add a data validation list to the merged cells
            // -------------------------------------------------
            // Create a new validation rule
            int validationIndex = worksheet.Validations.Add();
            Validation validation = worksheet.Validations[validationIndex];
            validation.Type = ValidationType.List;
            // Define the list items (comma‑separated)
            validation.Formula1 = "Option1,Option2,Option3";

            // Apply the validation to the merged range D2:E2
            CellArea area = new CellArea
            {
                StartRow = 1,      // row 2
                StartColumn = 3,   // column D
                EndRow = 1,        // row 2
                EndColumn = 4      // column E
            };
            validation.AddArea(area);

            // -------------------------------------------------
            // Save the workbook with merged areas validation enabled
            // -------------------------------------------------
            // Enable merging of validation areas during save (optional but recommended)
            XlsSaveOptions saveOptions = new XlsSaveOptions
            {
                MergeAreas = true,
                ValidateMergedAreas = true
            };
            workbook.Save("MergedAndValidated.xlsx", saveOptions);
        }
    }
}