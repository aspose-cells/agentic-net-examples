// Title: C# – List All PivotFields and Log Each PivotItem’s Absolute Position with Aspose.Cells
// Description: This Aspose.Cells for .NET example creates a workbook, builds a simple pivot table, refreshes and calculates it, then iterates through Row, Column, Page, and Data fields. For each field it safely logs the field name and, when available, every PivotItem’s name together with its absolute Position index – a handy debugging technique before saving the workbook.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | PivotField | PivotItem | absolute position | item index | debug pivot table | list pivot items | iterate pivot fields | Aspose.Cells example | GitHub Aspose.Cells | Excel pivot debugging | retrieve pivot item position
// Common Searches: Aspose.Cells get PivotItem position C# | list all PivotFields in a pivot table Aspose.Cells | debug pivot table items Aspose.Cells .NET | how to log pivot item absolute index using Aspose.Cells | iterate pivot fields and items Aspose.Cells example
// Developer Intent: The developer needs to enumerate every PivotField in a pivot table and output each PivotItem’s name and absolute position for troubleshooting.
// Use Cases: Verify that RefreshData and CalculateData generate the expected PivotItems after modifying source data. | Identify missing or out‑of‑order items when programmatically adjusting pivot fields. | Provide detailed logs for custom grouping, sorting, or filtering logic in a pivot table. | Create diagnostic output for automated tests that validate pivot table structure.
// AI Prompts: Write C# code that loops through all PivotFields in an Aspose.Cells pivot table and prints each PivotItem’s Name and Position, handling null collections. | Generate a helper method for Aspose.Cells that logs a warning when a PivotField has no PivotItems and returns the list of positions. | Explain the relationship between a PivotItem’s Position property and its absolute index in the pivot cache for Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableDemo
{
    // This Aspose.Cells for .NET example creates a workbook, builds a simple pivot table, refreshes and calculates it, then iterates through Row, Column, Page, and Data fields. For each field it safely logs the field name and, when available, every PivotItem’s name together with its absolute Position index – a handy debugging technique before saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["A4"].PutValue("A");
                worksheet.Cells["B4"].PutValue(30);

                // Add a pivot table based on the data range
                int pivotIndex = worksheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

                // Refresh and calculate the pivot table to ensure items are generated
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Log items of each field type, guarding against null PivotItems
                foreach (PivotField field in pivotTable.RowFields)
                    LogPivotFieldItems(field);

                foreach (PivotField field in pivotTable.ColumnFields)
                    LogPivotFieldItems(field);

                foreach (PivotField field in pivotTable.PageFields)
                    LogPivotFieldItems(field);

                foreach (PivotField field in pivotTable.DataFields)
                    LogPivotFieldItems(field);

                // Save the workbook
                workbook.Save("PivotDebugOutput.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Helper method to log each PivotItem's name and absolute position
        static void LogPivotFieldItems(PivotField field)
        {
            if (field == null) return;

            Console.WriteLine($"PivotField: {field.Name}");

            // Some field types (e.g., Data fields) may not have PivotItems
            if (field.PivotItems == null) return;

            foreach (PivotItem item in field.PivotItems)
            {
                // Position provides the absolute index of the item among all items
                Console.WriteLine($"  Item Name: {item.Name}, Absolute Position: {item.Position}");
            }
        }
    }
}
