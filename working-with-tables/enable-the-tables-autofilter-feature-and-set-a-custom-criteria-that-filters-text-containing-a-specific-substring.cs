// Title: C# Aspose.Cells Example: Enable Table Auto‑Filter and Apply a Contains Text Criterion
// Description: Creates a workbook, defines a ListObject over A1:B5, turns on the table's auto‑filter, applies a custom "Contains" filter on the Product column for the substring "Ap", refreshes the filter, and saves the file as TableAutoFilterContains.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Excel table auto filter | ListObject | custom filter | Contains operator | substring filter | sample code | GitHub example
// Common Searches: Aspose.Cells enable auto filter C# | filter Aspose.Cells table by text | ListObject custom contains filter .NET | apply substring criteria to Excel table using Aspose | C# code sample for table auto‑filter Aspose.Cells
// Developer Intent: Turn on a table’s auto‑filter and show only rows whose first column contains a specified substring.
// Use Cases: Display products whose names include a keyword without manual Excel interaction. | Generate reports that automatically hide non‑matching rows before exporting. | Programmatically switch filter text to create different data views in the same workbook.
// AI Prompts: Generate C# code with Aspose.Cells to add a ListObject, enable auto‑filter, and apply a Contains filter for "Ap" on column 0. | Show how to change the filter to StartsWith or EndsWith for a table column in Aspose.Cells .NET. | Explain how to refresh the auto‑filter after modifying custom filter criteria in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableAutoFilter
{
    // Creates a workbook, defines a ListObject over A1:B5, turns on the table's auto‑filter, applies a custom "Contains" filter on the Product column for the substring "Ap", refreshes the filter, and saves the file as TableAutoFilterContains.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (header + rows)
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue("Fruit");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue("Fruit");
            sheet.Cells["A4"].PutValue("Carrot");
            sheet.Cells["B4"].PutValue("Vegetable");
            sheet.Cells["A5"].PutValue("Apricot");
            sheet.Cells["B5"].PutValue("Fruit");

            // Create a ListObject (table) that covers the data range A1:B5
            // Parameters: startRow, startColumn, endRow, endColumn, hasHeaders
            int tableIndex = sheet.ListObjects.Add(0, 0, 4, 1, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Enable auto‑filter for the table
            table.HasAutoFilter = true;

            // Apply a custom filter on the first column (Product) to show rows
            // that contain the substring "Ap"
            table.AutoFilter.Custom(0, FilterOperatorType.Contains, "Ap");
            table.AutoFilter.Refresh();

            // Save the workbook
            workbook.Save("TableAutoFilterContains.xlsx");
        }
    }
}
