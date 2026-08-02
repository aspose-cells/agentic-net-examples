// Title: Add a PivotTable Timeline and Export to PDF with Aspose.Cells for .NET
// Description: Load an Excel workbook, create a PivotTable from Date and Sales columns, attach a Timeline control, save the updated file, and convert it to PDF using Aspose.Cells and ConversionUtility, with robust error handling.
// Keywords: Aspose.Cells | C# timeline control | PivotTable timeline | Excel to PDF conversion | export workbook as PDF | .NET Excel automation | sales report PDF | ConversionUtility | add timeline programmatically | pivot table C# example
// Common Searches: how to add a timeline to a pivot table using Aspose.Cells | asp.net convert excel with timeline to pdf | c# create pivot table and timeline Aspose.Cells | export excel workbook with timeline to pdf | aspose.cells timeline control example
// Developer Intent: Create a PivotTable, link a Timeline control, and generate a PDF from the modified worksheet.
// Use Cases: Produce a sales report PDF that lets readers filter data by date via an interactive timeline. | Automate monthly workbook processing: add a timeline for date selection and output a ready‑to‑share PDF. | Prepare a temporary Excel file with a timeline for downstream workflows before final PDF conversion.
// AI Prompts: Generate C# code with Aspose.Cells that builds a PivotTable from columns A and B, adds a Timeline linked to the Date field, and saves the result as PDF. | Explain how to calculate the last data row in a worksheet and construct the source range for a PivotTable in Aspose.Cells. | Provide best‑practice error handling when inserting a Timeline control and converting an Excel file to PDF using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Utility;

namespace AsposeCellsTimelineToPdf
{
    // Load an Excel workbook, create a PivotTable from Date and Sales columns, attach a Timeline control, save the updated file, and convert it to PDF using Aspose.Cells and ConversionUtility, with robust error handling.
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths – adjust as needed
                string inputPath = "SalesData.xlsx";                 // Existing workbook with sales data
                string tempPath = "SalesData_WithTimeline.xlsx";
                string pdfPath = "SalesReport.pdf";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // 1. Create a PivotTable based on the sales data.
                //    Assume the data is in columns A (Date) and B (Sales) with headers in row 1.
                // -------------------------------------------------
                // Determine the last row that contains data in column A (Date)
                int lastDataRow = sheet.Cells.GetLastDataRow(0); // 0‑based index
                if (lastDataRow < 1) // No data rows beyond header
                {
                    Console.WriteLine("The source worksheet does not contain data rows.");
                    return;
                }

                // Build the source range string (e.g., "A1:B5")
                string sourceRange = $"A1:B{lastDataRow + 1}";

                // Add the PivotTable to the worksheet (placed starting at cell D2)
                int pivotIndex = sheet.PivotTables.Add(sourceRange, "D2", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Configure the PivotTable: Date as row field, Sales as data field
                // Use column indexes to avoid mismatched header names
                pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Column A – Date
                pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Column B – Sales

                // Refresh and calculate the PivotTable so it contains data
                pivot.RefreshData();
                pivot.CalculateData();

                // -------------------------------------------------
                // 2. Add a Timeline control linked to the PivotTable.
                //    Place the Timeline at cell F1 (row 0, column 5).
                // -------------------------------------------------
                // Retrieve the actual field name used for the Date row field
                string dateFieldName = pivot.RowFields[0].Name;

                // Add the Timeline; wrap in try‑catch to handle potential issues
                try
                {
                    sheet.Timelines.Add(pivot, 0, 5, dateFieldName);
                }
                catch (Exception tlEx)
                {
                    Console.WriteLine($"Failed to add Timeline: {tlEx.Message}");
                    // Continue without Timeline if not critical
                }

                // -------------------------------------------------
                // 3. Save the workbook (now containing the Timeline) to a temporary file.
                // -------------------------------------------------
                workbook.Save(tempPath);

                // -------------------------------------------------
                // 4. Convert the temporary workbook to PDF.
                //    Using ConversionUtility which internally loads the source file and saves as PDF.
                // -------------------------------------------------
                // Ensure the temporary file was created before conversion
                if (File.Exists(tempPath))
                {
                    ConversionUtility.Convert(tempPath, pdfPath);
                    Console.WriteLine("Timeline added and workbook exported to PDF successfully.");
                }
                else
                {
                    Console.WriteLine("Temporary workbook file was not created; PDF conversion skipped.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
