// Title: Repeat Row 1 as Print Title (Column Headings) on Every Page with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds header values in row 1, fills rows 2‑50 with sample data, sets PageSetup.PrintTitleRows = "$1:$1" so the first row repeats on each printed page, optionally defines the print area, and saves the file as PrintTitleRowsDemo.xlsx.
// Keywords: Aspose.Cells PrintTitleRows | C# repeat header row print | Aspose.Cells page setup repeat rows | Aspose.Cells set print titles | Aspose.Cells define print area | Aspose.Cells .NET example | Excel repeat first row printing | Aspose.Cells workbook printing
// Common Searches: How to repeat the first row on every printed page using Aspose.Cells .NET | Aspose.Cells C# set PrintTitleRows property | Define print area and repeat header row in Aspose.Cells | Aspose.Cells page setup repeat rows example | Print titles in Excel with Aspose.Cells C#
// Developer Intent: Configure a worksheet so that the first row appears as a print title on each page of the printed document.
// Use Cases: Multi‑page reports where column headings must appear on every printed sheet. | Invoices or statements that span several pages and need the header row repeated for clarity. | Exporting data grids to Excel with a fixed print area and repeating titles for professional printing.
// AI Prompts: Show how to set multiple rows as print titles using Aspose.Cells for .NET. | Provide a C# example that configures both PrintTitleRows and PrintTitleColumns together with a custom print area. | Explain how to combine page orientation, scaling, and repeating titles in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsPrintTitleDemo
{
    // Creates a workbook, adds header values in row 1, fills rows 2‑50 with sample data, sets PageSetup.PrintTitleRows = "$1:$1" so the first row repeats on each printed page, optionally defines the print area, and saves the file as PrintTitleRowsDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample header in the first row (optional, just for demonstration)
            worksheet.Cells["A1"].PutValue("Header 1");
            worksheet.Cells["B1"].PutValue("Header 2");
            worksheet.Cells["C1"].PutValue("Header 3");

            // Populate some data below the header
            for (int i = 2; i <= 50; i++)
            {
                worksheet.Cells[$"A{i}"].PutValue($"Data A{i}");
                worksheet.Cells[$"B{i}"].PutValue($"Data B{i}");
                worksheet.Cells[$"C{i}"].PutValue($"Data C{i}");
            }

            // Access the page setup of the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // Set the first row to repeat on each printed page
            pageSetup.PrintTitleRows = "$1:$1";

            // Optionally define a print area (e.g., all used cells)
            pageSetup.PrintArea = "A1:C50";

            // Save the workbook to a file
            workbook.Save("PrintTitleRowsDemo.xlsx");
        }
    }
}
