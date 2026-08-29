// Title: Add a PivotTable timeline to an Excel worksheet and export it as PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, populates it with date‑sales data, builds a pivot table, inserts a timeline linked to the Date field, and saves the result as a PDF with Aspose.Cells. | Write a C# program that adds a timeline control to an existing pivot table in an Excel sheet and then converts the sheet to a PDF file using Aspose.Cells ConversionUtility.
// Common Searches: how to programmatically add a timeline to a pivot table in Aspose.Cells C# | Aspose.Cells convert Excel workbook with timeline control to PDF | C# example of creating a sales pivot table with a timeline and exporting to PDF | using Aspose.Cells Timeline API to filter pivot data before PDF conversion | sample code for adding timeline control to Excel file and saving as PDF in .NET
// Tags: Aspose.Cells add timeline to pivot table | export Excel with timeline to PDF C# | pivot table timeline control Aspose.Cells | convert workbook to PDF using ConversionUtility | sales data pivot timeline example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Utility;

namespace AsposeCellsTimelineToPdf
{
    // The program creates a new workbook, fills it with sample sales dates and amounts, builds a pivot table, attaches a timeline control to the Date field, saves the workbook as an Excel file, and then converts the file to PDF using Aspose.Cells' ConversionUtility.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate worksheet with sample sales data (Date and Sales columns)
                cells["A1"].PutValue("Date");
                cells["B1"].PutValue("Sales");

                cells["A2"].PutValue(new DateTime(2023, 1, 1));
                cells["B2"].PutValue(1500);

                cells["A3"].PutValue(new DateTime(2023, 2, 1));
                cells["B3"].PutValue(2300);

                cells["A4"].PutValue(new DateTime(2023, 3, 1));
                cells["B4"].PutValue(1800);

                cells["A5"].PutValue(new DateTime(2023, 4, 1));
                cells["B5"].PutValue(2100);

                // Create a pivot table based on the sales data
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D2", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh and calculate the pivot table data
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a timeline control linked to the pivot table (placed at cell F2)
                // Use zero‑based row/column indices: row 1 (second row), column 5 (column F)
                sheet.Timelines.Add(pivot, 1, 5, "Date");

                // Save the workbook to a temporary Excel file
                string tempExcelPath = "SalesTimeline.xlsx";
                workbook.Save(tempExcelPath);

                // Convert the Excel file (with the timeline) to PDF using the provided conversion rule
                string pdfPath = "SalesTimeline.pdf";

                // Ensure the source file exists before conversion
                if (File.Exists(tempExcelPath))
                {
                    ConversionUtility.Convert(tempExcelPath, pdfPath);
                    Console.WriteLine("Timeline added and workbook exported to PDF successfully.");
                }
                else
                {
                    Console.WriteLine($"Error: The file '{tempExcelPath}' was not found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
