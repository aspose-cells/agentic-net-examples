// Title: Aspose.Cells for .NET – Create an Excel Table and Apply List‑Type Validation Using a Table Column (C#)
// Description: Demonstrates how to generate a workbook, add a ListObject (Excel table) named ItemsTable, populate its Item column, define a validation range (C2:C10), set a List‑type data validation with an in‑cell dropdown, and reference the table column via a structured reference (=ItemsTable[Item]) before saving the file.
// Keywords: Aspose.Cells C# | Excel table ListObject | data validation list from table column | structured reference validation | in‑cell dropdown Aspose | .NET Excel dropdown list | create table validation Aspose.Cells
// Common Searches: Aspose.Cells create table with ListObject C# | list validation referencing table column Aspose.Cells | structured reference for data validation .NET | Excel dropdown list from table column using Aspose | how to add data validation list to range in Aspose.Cells
// Developer Intent: Create an Excel table and configure a list‑type data validation that pulls its allowed values from a column of that table using Aspose.Cells for .NET.
// Use Cases: Build a template where users select items from a master list that automatically updates when the table changes. | Generate workbooks with synchronized dropdowns in column C that reflect additions or deletions in the ItemsTable. | Apply a reusable validation rule across multiple worksheets by referencing a single table column.
// AI Prompts: Show how to apply the same table‑based validation to several non‑contiguous ranges in a worksheet. | Provide code to save the workbook to a MemoryStream while preserving the table and validation. | Explain how to use a named range instead of a ListObject as the source for a list validation in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

namespace AsposeCellsTableValidationDemo
{
    // Demonstrates how to generate a workbook, add a ListObject (Excel table) named ItemsTable, populate its Item column, define a validation range (C2:C10), set a List‑type data validation with an in‑cell dropdown, and reference the table column via a structured reference (=ItemsTable[Item]) before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ---------- Create source table ----------
                // Fill source data (column A)
                sheet.Cells["A1"].PutValue("Item");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Cherry");

                // Add a ListObject (Excel table) covering the range A1:A4
                int tableIdx = sheet.ListObjects.Add("A1", "A4", true);
                ListObject sourceTable = sheet.ListObjects[tableIdx];
                sourceTable.DisplayName = "ItemsTable"; // optional friendly name

                // ---------- Create validation referencing the table column ----------
                // Define the area where validation will be applied (C2:C10)
                CellArea validationArea = CellArea.CreateCellArea(1, 2, 9, 2); // rows 2-10, column C (index 2)

                // Add a new validation to the worksheet for the defined area
                int validationIdx = sheet.Validations.Add(validationArea);
                Validation validation = sheet.Validations[validationIdx];

                // Set validation type to List and enable the in‑cell dropdown
                validation.Type = ValidationType.List;
                validation.InCellDropDown = true;

                // Use a structured reference to the table column as the source list
                // Syntax: =TableName[ColumnName]
                validation.Formula1 = $"={sourceTable.DisplayName}[Item]";

                // ---------- Save the workbook ----------
                string outputPath = "TableValidationDemo.xlsx";

                // Ensure the directory exists before saving
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
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
