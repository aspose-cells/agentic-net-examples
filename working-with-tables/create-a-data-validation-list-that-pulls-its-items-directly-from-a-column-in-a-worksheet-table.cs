// Title: Create a data‑validation drop‑down in a cell that references a ListObject column using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells to define a ListObject table and attach a list‑type validation to cell C2 using the structured reference =ItemsTable[Item]. | Show how to configure Validation properties such as InCellDropDown, ErrorTitle, and ErrorMessage for a table‑based drop‑down list in an Excel workbook.
// Common Searches: aspnet aspocells add data validation list from table column c# | how to use structured reference for data validation in Aspose.Cells | C# Aspose.Cells create drop down list that reads values from ListObject | Aspose.Cells validation formula referencing worksheet table column | example of adding data validation to a specific cell using Aspose.Cells .NET
// Tags: Aspose.Cells list validation from ListObject column | C# add data validation list to cell | structured reference formula in Aspose.Cells | Excel table column drop‑down using Aspose.Cells | Aspose.Cells create ListObject table

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example creates a new workbook, defines a one‑column ListObject named ItemsTable, and adds a list‑type data validation to cell C2. The validation uses the structured reference =ItemsTable[Item] to populate the drop‑down, includes an error message, and saves the file as DataValidationFromTable.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate data that will become the table (column A)
            sheet.Cells["A1"].PutValue("Item");   // Header
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["A5"].PutValue("Date");

            // Define a table (ListObject) over the range A1:A5
            int firstRow = 0;      // zero‑based index for row 1
            int firstColumn = 0;   // zero‑based index for column A
            int totalRows = 5;
            int totalColumns = 1;
            // hasHeaders = true because the first row contains column name
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn,
                                                   firstRow + totalRows - 1,
                                                   firstColumn + totalColumns - 1,
                                                   true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "ItemsTable";   // Structured table name
            table.ShowTotals = false;           // No totals row

            // Add data‑validation list to cell C2 that pulls values from the table column
            // Structured reference syntax: =ItemsTable[Item]
            ValidationCollection validations = sheet.Validations;

            // Define the cell area for C2 (row index 1, column index 2)
            CellArea area = new CellArea
            {
                StartRow = 1,
                StartColumn = 2,
                EndRow = 1,
                EndColumn = 2
            };
            int validationIndex = validations.Add(area);
            Validation validation = validations[validationIndex];
            validation.Type = ValidationType.List;
            validation.Formula1 = "=ItemsTable[Item]";   // Reference to the table column
            validation.InCellDropDown = true;
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid selection";
            validation.ErrorMessage = "Please select a value from the list.";

            // Ensure output directory exists
            string outputPath = "DataValidationFromTable.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
