// Title: Create an Excel workbook and populate a worksheet with monthly sales figures using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to instantiate a Workbook, add a worksheet named "SalesData", write a header row, loop through month and sales arrays to fill columns A and B, and save the file as an XLSX. | Generate a method that accepts string[] months and double[] sales, writes them into a worksheet with headers using Aspose.Cells, and returns the path of the saved workbook.
// Common Searches: Aspose.Cells C# example for writing monthly sales numbers to an Excel sheet | How to programmatically add a header row and data rows to an XLSX file with Aspose.Cells | C# populate Excel columns from string and double arrays using Aspose.Cells | Saving a workbook as SalesData.xlsx after filling data with Aspose.Cells
// Tags: Aspose.Cells create workbook and write data | populate worksheet cells from arrays C# | save workbook as XLSX using Aspose.Cells | add header row to Excel sheet Aspose.Cells | write month and sales values to Excel with Aspose

using Aspose.Cells;
using System;

// The program creates a new Workbook, adds a worksheet named "SalesData", writes a header row and six rows of month and sales figures into columns A and B, and saves the result as "SalesData.xlsx".
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "SalesData";

        // Populate header row
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["B1"].PutValue("Sales");

        // Sample sales data
        string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
        double[] sales = { 12000, 15000, 13000, 17000, 16000, 18000 };

        // Fill data into cells
        for (int i = 0; i < months.Length; i++)
        {
            sheet.Cells[i + 2, 0].PutValue(months[i]); // Column A (Month)
            sheet.Cells[i + 2, 1].PutValue(sales[i]); // Column B (Sales)
        }

        // Save the workbook to a file
        workbook.Save("SalesData.xlsx", SaveFormat.Xlsx);
    }
}
