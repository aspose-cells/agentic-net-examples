// Title: C# – Apply Top 10 AutoFilter on Sales Column with Aspose.Cells for .NET
// Description: Creates a workbook, adds product and sales data, defines an AutoFilter range, applies FilterTop10 to the Sales column to keep the highest‑selling items, refreshes the filter, and saves the result as Top10SalesFilter.xlsx.
// Keywords: Aspose.Cells | C# | AutoFilter | FilterTop10 | Top 10 filter | Excel sales report | highest selling products | filter rows by value | example code
// Common Searches: Aspose.Cells top 10 filter C# example | FilterTop10 method usage Aspose.Cells | Show top selling items in Excel with Aspose | C# auto filter top 10 rows | How to filter sales column by top values using Aspose.Cells
// Developer Intent: Create an Excel workbook and filter the Sales column so that only the top‑10 products by sales are visible.
// Use Cases: Produce a concise sales dashboard that automatically lists the best‑selling products. | Enable analysts to quickly isolate high‑performing items without manual sorting. | Export the filtered top‑10 list to a separate file for stakeholder distribution.
// AI Prompts: Generate C# code with Aspose.Cells that applies a Top 5 filter on a user‑specified column and saves the filtered workbook. | Explain each parameter of the FilterTop10 method and show how to switch between count‑based and percent‑based filtering. | Provide a combined example that uses AutoFilter to select the top rows and conditional formatting to highlight them.

using System;
using Aspose.Cells;

namespace AsposeCellsTop10FilterDemo
{
    // Creates a workbook, adds product and sales data, defines an AutoFilter range, applies FilterTop10 to the Sales column to keep the highest‑selling items, refreshes the filter, and saves the result as Top10SalesFilter.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data: Product names and their sales figures
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Sales");

            string[] products = { "Apple", "Banana", "Orange", "Grapes", "Mango", "Pineapple", "Strawberry" };
            int[] sales = { 150, 85, 230, 120, 190, 75, 200 };

            for (int i = 0; i < products.Length; i++)
            {
                cells[i + 1, 0].PutValue(products[i]);   // Column A: Product
                cells[i + 1, 1].PutValue(sales[i]);      // Column B: Sales
            }

            // Define the autofilter range covering both columns and all rows with data
            sheet.AutoFilter.Range = "A1:B8";

            // Apply a Top 10 filter on the Sales column (field index 1)
            // Show the top 10 items by count (not percent)
            sheet.AutoFilter.FilterTop10(fieldIndex: 1, isTop: true, isPercent: false, itemCount: 10);

            // Refresh the filter to hide rows that do not meet the criteria
            sheet.AutoFilter.Refresh();

            // Save the workbook to a file
            workbook.Save("Top10SalesFilter.xlsx");
        }
    }
}
