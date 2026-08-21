// Title: C# Aspose.Cells Top 10 AutoFilter for Highest‑Selling Products
// Description: Creates a workbook, populates product and sales columns, defines an AutoFilter range, applies a Top 10 filter on the Sales column to keep the ten largest values, refreshes the view, and saves the file as Top10Products.xlsx.
// Keywords: Aspose.Cells | C# | .NET | AutoFilter | Top 10 filter | highest sales | product ranking | Excel automation | filter by top values
// Common Searches: Aspose.Cells Top 10 filter C# example | How to show top selling items with Aspose.Cells | C# AutoFilter highest sales Aspose | Apply Top10 filter on Excel column using Aspose.Cells | Filter top 10 rows by value in .NET
// Developer Intent: Generate an Excel workbook and use Aspose.Cells to display only the ten products with the greatest sales figures.
// Use Cases: Quickly produce a sales dashboard that highlights the best‑performing products. | Create a reusable report template that automatically filters to the top‑10 items for monthly reviews. | Export a pre‑filtered list of top sellers to downstream systems or BI tools.
// AI Prompts: Write C# code with Aspose.Cells to apply a Top 10 AutoFilter on column B and save the workbook. | Modify the example to use a percentage‑based Top 10 filter instead of a count‑based filter. | Add conditional formatting that colors rows meeting the Top 10 criteria.

using System;
using Aspose.Cells;

namespace AsposeCellsTop10FilterDemo
{
    // Creates a workbook, populates product and sales columns, defines an AutoFilter range, applies a Top 10 filter on the Sales column to keep the ten largest values, refreshes the view, and saves the file as Top10Products.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add header row
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Sales");

            // Sample product and sales data
            string[] products = { "Apple", "Banana", "Orange", "Grapes", "Mango", "Pineapple", "Strawberry", "Blueberry", "Kiwi", "Peach", "Cherry", "Lemon", "Watermelon", "Papaya", "Guava" };
            int[] sales = { 120, 85, 150, 60, 200, 95, 130, 70, 55, 110, 90, 65, 180, 75, 50 };

            // Populate the worksheet with data
            for (int i = 0; i < products.Length; i++)
            {
                cells[i + 1, 0].PutValue(products[i]); // Column A (Product)
                cells[i + 1, 1].PutValue(sales[i]);    // Column B (Sales)
            }

            // Define the autofilter range (including header)
            int lastRow = products.Length + 1; // +1 for header row
            sheet.AutoFilter.Range = $"A1:B{lastRow}";

            // Apply a Top 10 filter on the Sales column (field index 1)
            // isTop = true (show highest), isPercent = false (use count), itemCount = 10
            sheet.AutoFilter.FilterTop10(fieldIndex: 1, isTop: true, isPercent: false, itemCount: 10);

            // Refresh the filter to hide rows that do not meet the criteria
            sheet.AutoFilter.Refresh();

            // Save the workbook
            workbook.Save("Top10Products.xlsx");
        }
    }
}
