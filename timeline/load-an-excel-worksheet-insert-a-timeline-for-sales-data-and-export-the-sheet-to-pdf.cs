// Title: C# – Add a Timeline to a Pivot Table and Export Excel to PDF with Aspose.Cells
// Description: Creates a sample sales workbook if missing, builds a pivot table on Ship Date and Sales, attaches a Timeline control to the pivot, saves the updated file, and converts it to a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells timeline C# | add timeline to pivot table .NET | export Excel to PDF Aspose | pivot table with timeline example | C# Excel PDF conversion Aspose.Cells | timeline control Excel API | GitHub Aspose.Cells sample
// Common Searches: how to add a timeline to a pivot table using Aspose.Cells | convert Excel workbook with timeline to PDF C# | Aspose.Cells example for timeline control | C# code to create pivot table and timeline | export Excel with timeline to PDF Aspose
// Developer Intent: Insert a Timeline linked to a Pivot Table in an Excel sheet and generate a PDF report programmatically.
// Use Cases: Automated sales dashboards that let users filter data via a timeline before publishing to PDF. | Batch processing of Excel files to add interactive timelines and produce printable reports. | Generating PDF snapshots of workbooks that include date‑driven visual controls for stakeholder review.
// AI Prompts: Provide C# code that creates a pivot table, adds a timeline control, and exports the workbook to PDF with Aspose.Cells. | Show how to handle a missing source Excel file by generating sample sales data before adding a timeline and converting to PDF. | Explain the steps to link a Timeline to the 'Ship Date' field of a pivot table and convert the result to PDF using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Utility;

// Creates a sample sales workbook if missing, builds a pivot table on Ship Date and Sales, attaches a Timeline control to the pivot, saves the updated file, and converts it to a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "SalesData.xlsx";
            const string tempFile = "SalesData_WithTimeline.xlsx";
            const string pdfFile = "SalesReport.pdf";

            // Ensure the source workbook exists; create a simple one if it does not.
            if (!File.Exists(inputFile))
            {
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                // Add headers
                ws.Cells["A1"].PutValue("Ship Date");
                ws.Cells["B1"].PutValue("Sales");
                // Add sample data
                ws.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
                ws.Cells["B2"].PutValue(1200);
                ws.Cells["A3"].PutValue(new DateTime(2023, 1, 2));
                ws.Cells["B3"].PutValue(1500);
                ws.Cells["A4"].PutValue(new DateTime(2023, 1, 3));
                ws.Cells["B4"].PutValue(800);
                ws.Cells["A5"].PutValue(new DateTime(2023, 1, 4));
                ws.Cells["B5"].PutValue(950);
                wb.Save(inputFile);
            }

            // Load the workbook containing sales data.
            Workbook workbook = new Workbook(inputFile);
            Worksheet sheet = workbook.Worksheets[0];

            // Create a pivot table from the data range A1:B5 and place it at D1.
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure pivot fields.
            pivot.AddFieldToArea(PivotFieldType.Row, "Ship Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table.
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline control linked to the pivot table at cell F1.
            sheet.Timelines.Add(pivot, "F1", "Ship Date");

            // Save the workbook with the Timeline to a temporary file.
            workbook.Save(tempFile);

            // Convert the workbook to PDF.
            ConversionUtility.Convert(tempFile, pdfFile);

            Console.WriteLine($"PDF generated successfully: {pdfFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
