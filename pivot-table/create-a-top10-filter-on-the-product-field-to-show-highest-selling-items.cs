// Title: Apply a Top 10 AutoFilter to the Sales column in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates an Excel workbook with Product and Sales columns, defines an AutoFilter range, and uses Aspose.Cells FilterTop10 to display only the top 10 sales rows. | Show how to modify the Top 10 filter to a percentage‑based filter (e.g., top 10 % of sales) with Aspose.Cells AutoFilter in C#. | Explain the steps to refresh the AutoFilter after calling FilterTop10 and save the filtered workbook to a file using Aspose.Cells.
// Common Searches: asp.net c# aspocells how to filter top 10 rows by sales column | example code for applying top 10 autofilter on Excel sheet using Aspose.Cells | filter highest selling products in Excel with Aspose.Cells C# | using FilterTop10 method in Aspose.Cells to show top sales items | Aspose.Cells AutoFilter range and top 10 filter tutorial
// Tags: Aspose.Cells FilterTop10 API usage C# | Excel AutoFilter range definition Aspose.Cells | Top 10 sales filter Aspose.Cells C# | percentage based top filter Aspose.Cells | refresh AutoFilter after FilterTop10 Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsTop10FilterDemo
{
    // The sample creates a new workbook, fills columns A and B with product names and sales figures, sets an AutoFilter covering the data range, applies Aspose.Cells' FilterTop10 to keep only the ten highest‑selling rows, refreshes the filter to hide other rows, and saves the result as Top10SalesFilter.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: Product names and Sales values
            // Header row
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");

            // Sample data rows
            string[] products = { "Apple", "Banana", "Orange", "Grapes", "Mango", "Pineapple", "Strawberry", "Kiwi", "Peach", "Cherry", "Lemon", "Watermelon" };
            int[] sales = { 120, 85, 150, 60, 200, 95, 110, 70, 130, 55, 40, 180 };

            for (int i = 0; i < products.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(products[i]);   // Column A: Product
                sheet.Cells[i + 1, 1].PutValue(sales[i]);     // Column B: Sales
            }

            // Define the autofilter range covering both columns (including header)
            sheet.AutoFilter.Range = "A1:B13";

            // Apply a Top 10 filter on the Sales column (field index 1)
            // Parameters: fieldIndex, isTop, isPercent, itemCount
            sheet.AutoFilter.FilterTop10(fieldIndex: 1, isTop: true, isPercent: false, itemCount: 10);

            // Refresh the filter to hide rows that do not meet the criteria
            sheet.AutoFilter.Refresh();

            // Save the workbook
            workbook.Save("Top10SalesFilter.xlsx");
        }
    }
}
