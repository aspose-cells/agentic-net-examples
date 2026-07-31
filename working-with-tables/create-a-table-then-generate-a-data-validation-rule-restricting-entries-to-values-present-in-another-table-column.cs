// Title: Aspose.Cells .NET – Create an Excel Table and Apply List‑Data Validation from a Table Column (C#)
// Description: Demonstrates how to build a source table (ItemsTable) from cells A1:A5, define a validation range (C2:C10), and attach a List‑type data validation that references the table column via Validation.Formula1 (e.g., =ItemsTable[Item]). The drop‑down list updates automatically with the table and the workbook is saved as TableValidationDemo.xlsx.
// Keywords: Aspose.Cells C# table creation | Excel ListObject data validation | Aspose.Cells Validation.Formula1 | C# generate drop‑down list from table | Aspose.Cells dynamic validation range | Excel table column reference | Aspose.Cells .NET example | C# workbook save Aspose
// Common Searches: Aspose.Cells set data validation list from table column | C# create Excel table and add drop‑down validation | How to reference ListObject column in Validation.Formula1 | Aspose.Cells example for list‑type validation | Apply data validation to multiple cells Aspose.Cells
// Developer Intent: Create a table and attach a list‑type validation that pulls its allowed values from the table column.
// Use Cases: Provide users with a consistent pick‑list sourced from a master table. | Build templates where dropdown options automatically reflect changes in the source table. | Enforce data integrity in reports by limiting entries to predefined table values.
// AI Prompts: Generate C# code with Aspose.Cells that creates a table from A1:A5 and adds a list validation to C2:C10 referencing the table column. | Explain the syntax for Validation.Formula1 when using a ListObject column in Aspose.Cells. | Show how to apply the same Validation object to several non‑contiguous ranges in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to build a source table (ItemsTable) from cells A1:A5, define a validation range (C2:C10), and attach a List‑type data validation that references the table column via Validation.Formula1 (e.g., =ItemsTable[Item]). The drop‑down list updates automatically with the table and the workbook is saved as TableValidationDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Source table (values to be used in validation) ----------
            // Header
            sheet.Cells["A1"].PutValue("Item");
            // Data rows
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["A5"].PutValue("Date");

            // Create a ListObject (Excel table) from the range A1:A5
            int tableIdx = sheet.ListObjects.Add("A1", "A5", true);
            ListObject sourceTable = sheet.ListObjects[tableIdx];
            sourceTable.DisplayName = "ItemsTable"; // optional friendly name

            // Ensure the first column has a name (taken from header cell)
            if (string.IsNullOrEmpty(sourceTable.ListColumns[0].Name))
            {
                sourceTable.ListColumns[0].Name = sheet.Cells["A1"].StringValue;
            }

            // ---------- Target column where validation will be applied ----------
            sheet.Cells["C1"].PutValue("Selection"); // header for the target column

            // Define the area (C2:C10) that will have the drop‑down validation
            CellArea validationArea = CellArea.CreateCellArea(1, 2, 9, 2); // rows 2‑10, column C (index 2)

            // Get a Validation object from the first cell of the area
            Validation validation = sheet.Cells["C2"].GetValidation();

            // Set validation type to List and point it to the source table column
            validation.Type = ValidationType.List;
            validation.Formula1 = $"={sourceTable.DisplayName}[{sourceTable.ListColumns[0].Name}]";

            // Show the in‑cell drop‑down list
            validation.InCellDropDown = true;

            // Apply the validation to the whole area
            validation.AddArea(validationArea);

            // Save the workbook
            string outputPath = "TableValidationDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
