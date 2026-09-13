// Title: How to repeat the header row and first column as print titles in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that sets row 1 and column A as repeating print titles and defines a print area for the worksheet. | Show the steps to configure PageSetup.PrintTitleRows and PageSetup.PrintTitleColumns in Aspose.Cells so the header and leftmost column appear on every printed page. | Create a sample workbook that populates data, applies repeating header and first column for printing, and saves the result as an XLSX file using Aspose.Cells.
// Common Searches: Aspose.Cells C# repeat first row on each printed page example | Set column A as print title in Aspose.Cells .NET workbook | Define print area and repeat titles with Aspose.Cells page setup C# | How to keep header and left column visible when printing Excel using Aspose.Cells | C# Aspose.Cells page setup repeat rows and columns for PDF export
// Tags: Aspose.Cells page setup repeat rows | Aspose.Cells page setup repeat columns | Aspose.Cells define worksheet print area | Aspose.Cells printing options for XLSX | Aspose.Cells C# repeat header row and left column

using System;
using Aspose.Cells;

namespace AsposeCellsPrintTitleExample
{
    // The program creates a new workbook, fills cells A1:D100 with sample data, configures the page setup so that row 1 and column A repeat on every printed page, defines the print area, and saves the file as RepeatedHeaderAndFirstColumn.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (header row and first column)
            // Header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Quantity");
            sheet.Cells["D1"].PutValue("Price");

            // First column as title and some data rows
            for (int i = 2; i <= 100; i++)
            {
                sheet.Cells[i - 1, 0].PutValue($"Item {i - 1}");
                sheet.Cells[i - 1, 1].PutValue($"Product {i - 1}");
                sheet.Cells[i - 1, 2].PutValue(i * 10);
                sheet.Cells[i - 1, 3].PutValue(i * 1.5);
            }

            // Set the header row (first row) to repeat on each printed page
            sheet.PageSetup.PrintTitleRows = "$1:$1";

            // Set the first column (A) to repeat on each printed page as a title
            sheet.PageSetup.PrintTitleColumns = "$A:$A";

            // Optional: define the print area to include all used cells
            sheet.PageSetup.PrintArea = "$A$1:$D$100";

            // Save the workbook to a file
            workbook.Save("RepeatedHeaderAndFirstColumn.xlsx", SaveFormat.Xlsx);
        }
    }
}
