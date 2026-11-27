// Title: Generate an Excel dropdown list from a worksheet table column using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create data validation list referencing a ListObject column | How to bind Excel dropdown to a table column with Aspose.Cells .NET | C# code for structured reference validation list in Aspose.Cells workbook | Add in‑cell dropdown from worksheet table using Aspose.Cells API | Programmatically set validation formula to =MyTable[Item] in Aspose.Cells
// Tags: worksheet table column data validation Aspose.Cells | structured reference =MyTable[Item] validation | ListObject creation and dropdown C# | Excel dropdown list from table column .NET | Aspose.Cells validation formula using table reference

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

namespace DataValidationFromTable
{
    // The example creates a new workbook, adds a single‑column ListObject named MyTable with sample items, defines a data validation rule for cell C2 that uses the structured reference "=MyTable[Item]" to display an in‑cell dropdown, and saves the file as DataValidationFromTable.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Get the first worksheet and name it.
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // ------------------------------------------------------------
                // 1. Populate a column with items that will serve as the list source.
                // ------------------------------------------------------------
                sheet.Cells["A1"].PutValue("Item");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["A5"].PutValue("Date");

                // ------------------------------------------------------------
                // 2. Convert the range into a worksheet table (ListObject).
                // ------------------------------------------------------------
                int firstRow = 0;          // zero‑based index (A1)
                int firstColumn = 0;       // column A
                int totalRows = 5;         // rows 1‑5 inclusive
                int totalColumns = 1;      // only one column

                // Add the table to the worksheet.
                int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn,
                    firstRow + totalRows - 1, firstColumn + totalColumns - 1, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "MyTable";

                // ------------------------------------------------------------
                // 3. Create a data validation list that references the table column.
                // ------------------------------------------------------------
                // Define the target cell (C2) where the dropdown will appear.
                CellArea targetArea = new CellArea
                {
                    StartRow = 1,   // Row 2 (zero‑based)
                    EndRow = 1,
                    StartColumn = 2, // Column C (zero‑based)
                    EndColumn = 2
                };

                // Add a new validation rule for the target area.
                int validationIndex = sheet.Validations.Add(targetArea);
                Validation validation = sheet.Validations[validationIndex];
                validation.Type = ValidationType.List;
                validation.InCellDropDown = true;   // Show dropdown arrow.
                validation.IgnoreBlank = true;
                validation.ShowError = true;

                // Use a structured reference to the table column.
                validation.Formula1 = "=MyTable[Item]";

                // Optional custom error message.
                validation.ErrorTitle = "Invalid Selection";
                validation.ErrorMessage = "Please select a value from the list.";

                // ------------------------------------------------------------
                // 4. Save the workbook.
                // ------------------------------------------------------------
                string outputPath = "DataValidationFromTable.xlsx";

                // Ensure the directory exists (if a path is provided).
                string directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
