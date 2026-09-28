// Title: Create a new Excel workbook and populate it with monthly sales data using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to create a workbook, add a header row for Month and three products, and fill the rows with sales figures from arrays. | Show a loop that iterates over month and sales arrays and writes the values into cells A2:D7 of an Aspose.Cells worksheet. | Demonstrate saving the completed worksheet as an XLSX file named SalesData.xlsx with Aspose.Cells.
// Common Searches: aspocells c# example populate worksheet with sales data from arrays | how to add header row and data rows to Excel using Aspose.Cells .NET | save workbook as xlsx using Aspose.Cells C# loop through arrays | populate cells A1:D7 with month and product sales using Aspose.Cells | create Excel file with monthly sales chart data programmatically in C#
// Tags: Aspose.Cells create workbook and fill cells | populate worksheet from multidimensional array Aspose.Cells | write header row Aspose.Cells C# | save workbook as XLSX Aspose.Cells | loop through arrays to write Excel cells Aspose.Cells | initialize sales data worksheet Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsSalesDataExample
{
    // // Creates a new Workbook, adds a header row for Month and three products, fills six rows with monthly sales figures from arrays, and saves the file as SalesData.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the Workbook() constructor rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate header row
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Product A");
            sheet.Cells["C1"].PutValue("Product B");
            sheet.Cells["D1"].PutValue("Product C");

            // Sample sales data (Month, Product A, Product B, Product C)
            string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
            double[,] sales = {
                { 12000, 15000, 13000 },
                { 14000, 16000, 15000 },
                { 13000, 17000, 14000 },
                { 15000, 18000, 16000 },
                { 16000, 19000, 17000 },
                { 17000, 20000, 18000 }
            };

            // Fill the worksheet with the sales data
            for (int i = 0; i < months.Length; i++)
            {
                int row = i + 2; // Data starts from row 2
                sheet.Cells[$"A{row}"].PutValue(months[i]);          // Month
                sheet.Cells[$"B{row}"].PutValue(sales[i, 0]);       // Product A
                sheet.Cells[$"C{row}"].PutValue(sales[i, 1]);       // Product B
                sheet.Cells[$"D{row}"].PutValue(sales[i, 2]);       // Product C
            }

            // Save the workbook (uses the Save(string, SaveFormat) rule)
            workbook.Save("SalesData.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook with sales data created successfully.");
        }
    }
}
