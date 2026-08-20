// Title: Create a data‑validation drop‑down list from an Excel table column using Aspose.Cells for .NET
// Description: This example builds a one‑column table (OptionsTable) in a new workbook, defines cell B1 as the validation area, adds a List‑type validation with InCellDropDown enabled, and sets Validation.Formula1 to the structured reference "OptionsTable[Options]" so the drop‑down items are sourced directly from the table column. The workbook is then saved as DataValidationFromTable.xlsx.
// Keywords: Aspose.Cells data validation | C# Excel drop‑down list | structured reference validation | ListObject validation Aspose | Excel table column data validation .NET | Aspose.Cells ListObject example | create drop‑down from table column
// Common Searches: Aspose.Cells create data validation list from table column | C# add drop‑down list using structured reference | How to bind Excel validation to ListObject column Aspose | Set in‑cell dropdown from table with Aspose.Cells | Reference table column in Validation.Formula1 C#
// Developer Intent: Add a drop‑down validation to a cell that pulls its items from a column of an Excel table.
// Use Cases: Provide users with a selectable list that updates automatically when the table data changes. | Apply the same table‑based validation to multiple cells or an entire range. | Maintain a single source of truth for option values by storing them in a ListObject.
// AI Prompts: Generate C# code that creates a data‑validation list referencing a ListObject column with Aspose.Cells. | Show how to apply a table‑based drop‑down validation to a range of cells in Aspose.Cells for .NET. | Explain the syntax of a structured reference for Validation.Formula1 in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// This example builds a one‑column table (OptionsTable) in a new workbook, defines cell B1 as the validation area, adds a List‑type validation with InCellDropDown enabled, and sets Validation.Formula1 to the structured reference "OptionsTable[Options]" so the drop‑down items are sourced directly from the table column. The workbook is then saved as DataValidationFromTable.xlsx.
class DataValidationFromTable
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate data that will become the table (column A)
            worksheet.Cells["A1"].PutValue("Options");   // Header
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["A4"].PutValue("Cherry");
            worksheet.Cells["A5"].PutValue("Date");

            // Add a ListObject (Excel table) that covers the range A1:A5
            int tableIdx = worksheet.ListObjects.Add("A1", "A5", true);
            ListObject table = worksheet.ListObjects[tableIdx];
            table.DisplayName = "OptionsTable"; // Optional: give the table a friendly name

            // Define the cell (B1) where the drop‑down list will appear
            CellArea validationArea = new CellArea
            {
                StartRow = 0,      // Row 1 (zero‑based)
                StartColumn = 1,   // Column B (zero‑based)
                EndRow = 0,
                EndColumn = 1
            };

            // Add a validation object to the worksheet for the defined area
            int validationIdx = worksheet.Validations.Add(validationArea);
            Validation validation = worksheet.Validations[validationIdx];

            // Configure the validation as a List that references the table column
            validation.Type = ValidationType.List;
            validation.InCellDropDown = true;
            // Structured reference to the table column "Options"
            validation.Formula1 = "OptionsTable[Options]";

            // Save the workbook
            workbook.Save("DataValidationFromTable.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
