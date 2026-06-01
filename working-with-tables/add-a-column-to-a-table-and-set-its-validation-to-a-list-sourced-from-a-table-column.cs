using System;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

namespace AsposeCellsTableValidationDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data (including the source list column)
                // Header row
                cells["A1"].PutValue("ID");
                cells["B1"].PutValue("Options");   // This column will be the source list
                cells["C1"].PutValue("Value");
                cells["D1"].PutValue("Choice");    // Column that will have validation

                // Data rows
                for (int i = 2; i <= 5; i++)
                {
                    cells[$"A{i}"].PutValue(i - 1);                 // ID
                    cells[$"B{i}"].PutValue($"Option{i - 1}");      // Options (source list)
                    cells[$"C{i}"].PutValue((i - 1) * 10);          // Some other value
                    // D column left empty – will receive validation
                }

                // Add a ListObject (table) that includes the new column D
                // Table range: A1:D5, hasHeaders = true
                int tableIndex = sheet.ListObjects.Add("A1", "D5", true);
                ListObject table = sheet.ListObjects[tableIndex];
                // Use DisplayName to set the table name (Name property not available in some versions)
                table.DisplayName = "SampleTable";

                // Define the area for validation: column D (index 3), rows 2‑5 (excluding header)
                CellArea validationArea = CellArea.CreateCellArea(1, 3, 4, 3); // rows 1‑4 zero‑based, column 3

                // Add validation to the worksheet
                ValidationCollection validations = sheet.Validations;
                int validationIndex = validations.Add(validationArea);
                Validation validation = validations[validationIndex];

                // Configure validation as a list sourced from the "Options" column (B2:B5)
                validation.Type = ValidationType.List;
                validation.InCellDropDown = true;
                validation.Formula1 = "$B$2:$B$5"; // absolute reference to source list range

                // Save the workbook
                string outputPath = "TableWithDropdownValidation.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}