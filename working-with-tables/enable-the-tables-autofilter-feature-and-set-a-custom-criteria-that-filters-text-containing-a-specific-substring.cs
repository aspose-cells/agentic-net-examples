// Title: Aspose.Cells C# – Enable Table AutoFilter and Apply a Contains Filter
// Description: Creates a workbook, adds a ListObject (Excel table) with product data, turns on the table's AutoFilter, and uses a custom Contains filter to show only rows where the "Product" column includes the substring "Apple". The workbook is then saved as an .xlsx file.
// Keywords: Aspose.Cells | C# | Excel table AutoFilter | ListObject filter | Contains filter | FilterOperatorType.Contains | .NET spreadsheet filtering | custom text filter Aspose
// Common Searches: Aspose.Cells enable auto filter on table | C# filter table rows by substring | Apply Contains filter to ListObject Aspose | How to use custom filter with Aspose.Cells | AutoFilter Custom Contains example .NET
// Developer Intent: Turn on AutoFilter for a worksheet table and restrict visible rows to those whose column value contains a specified text.
// Use Cases: Show only products that contain a keyword (e.g., "Apple") in a generated catalog. | Create dynamic reports that hide non‑matching rows without manual editing. | Prepare Excel files where end‑users can quickly filter data by typing a substring.
// AI Prompts: Generate C# code that adds an AutoFilter to an Aspose.Cells ListObject and filters column 0 for rows containing "Apple". | Explain how to use FilterOperatorType.Contains with Aspose.Cells to filter table data by a text fragment. | Provide a step‑by‑step example of applying multiple custom filters (Contains, DoesNotContain) on an Aspose.Cells table.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a ListObject (Excel table) with product data, turns on the table's AutoFilter, and uses a custom Contains filter to show only rows where the "Product" column includes the substring "Apple". The workbook is then saved as an .xlsx file.
    public class TableAutoFilterWithContainsDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the table (header + data rows)
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple iPhone");
            sheet.Cells["B2"].PutValue("Electronics");
            sheet.Cells["A3"].PutValue("Banana Bread");
            sheet.Cells["B3"].PutValue("Food");
            sheet.Cells["A4"].PutValue("Apple MacBook");
            sheet.Cells["B4"].PutValue("Electronics");
            sheet.Cells["A5"].PutValue("Cherry Pie");
            sheet.Cells["B5"].PutValue("Food");

            // Define the range of the table (including header row)
            int firstRow = 0;   // zero‑based index for row 1
            int firstCol = 0;   // column A
            int lastRow  = 4;   // row 5 (zero‑based)
            int lastCol  = 1;   // column B

            // Add a ListObject (Excel table) covering the data range
            int tableIndex = sheet.ListObjects.Add(firstRow, firstCol, lastRow, lastCol, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Enable auto‑filter for the table
            table.HasAutoFilter = true;

            // Get the AutoFilter object associated with the table
            AutoFilter autoFilter = table.AutoFilter;

            // Apply a custom filter that shows rows where the "Product" column (index 0)
            // contains the substring "Apple"
            autoFilter.Custom(0, FilterOperatorType.Contains, "Apple");

            // Refresh the filter to apply the criteria
            autoFilter.Refresh();

            // Save the workbook
            workbook.Save("TableAutoFilterContainsDemo.xlsx");
        }
    }
}
