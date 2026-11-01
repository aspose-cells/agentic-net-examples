// Title: Generate an Excel workbook with a 'Sales' sheet and insert product, region, and sales rows using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a new Workbook, adds a worksheet named "Sales", writes the headers "Product", "Region", "Sales", iterates over a collection of objects to fill rows, and saves the file as an .xlsx using Aspose.Cells. | Extend the example to calculate total sales, add a summary row with a formula, and apply bold formatting to the header and total rows using Aspose.Cells styling APIs.
// Common Searches: asp.net how to create an Excel file with a worksheet and fill it with data using Aspose.Cells | c# Aspose.Cells add worksheet named Sales and write rows from a list of objects | populate Excel cells from an array of anonymous objects with Aspose.Cells C# example | save workbook as .xlsx after inserting sales data with Aspose.Cells .NET
// Tags: create workbook and add worksheet Aspose.Cells | populate cells from C# collection Aspose.Cells | save workbook as xlsx Aspose.Cells | write header and data rows Aspose.Cells | insert sales records into Excel Aspose.Cells

using System;
using Aspose.Cells;

namespace SalesDataExample
{
    // // Creates a new Workbook, adds a "Sales" worksheet, writes header cells and sample sales records into columns A‑C, then saves the file as "SalesData.xlsx".
    class Program
    {
        static void Main(string[] args)
        {
            // Instantiate a new Workbook
            Workbook workbook = new Workbook();

            // Add a new worksheet named "Sales"
            Worksheet sheet = workbook.Worksheets[workbook.Worksheets.Add()];
            sheet.Name = "Sales";

            // Populate header row
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Region");
            sheet.Cells["C1"].PutValue("Sales");

            // Sample sales data
            var data = new[]
            {
                new { Product = "Laptop", Region = "North", Sales = 1500 },
                new { Product = "Smartphone", Region = "South", Sales = 2300 },
                new { Product = "Tablet", Region = "East", Sales = 1200 },
                new { Product = "Monitor", Region = "West", Sales = 800 }
            };

            // Fill data starting from row 2
            int rowIndex = 1; // zero‑based index (row 2 in Excel)
            foreach (var item in data)
            {
                sheet.Cells[rowIndex, 0].PutValue(item.Product);   // Column A
                sheet.Cells[rowIndex, 1].PutValue(item.Region);    // Column B
                sheet.Cells[rowIndex, 2].PutValue(item.Sales);     // Column C
                rowIndex++;
            }

            // Save the workbook to a file (adjust path as needed)
            workbook.Save("SalesData.xlsx", SaveFormat.Xlsx);
        }
    }
}
