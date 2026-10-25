// Title: Load SalesData.xlsx, retrieve PivotTable1, and export the worksheet to PDF using Aspose.Cells for .NET (C#)
// Common Searches: c# Aspose.Cells example to load an Excel file and export a specific worksheet as PDF | how to get a pivot table by name from a worksheet using Aspose.Cells in C# | export Excel workbook with PivotTable to PDF with Aspose.Cells .NET | Aspose.Cells C# code to verify file existence before opening workbook | generate PDF sales report from SalesData.xlsx using Aspose.Cells
// Tags: Aspose.Cells load workbook C# | Aspose.Cells get pivot table by name | Aspose.Cells save as PDF from worksheet | Aspose.Cells verify file exists before opening | Aspose.Cells generate sales report PDF

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // For PivotTable support

namespace TimelineToPdfExampleApp
{
    // The example checks for the presence of SalesData.xlsx, loads it into an Aspose.Cells Workbook, accesses the first worksheet, retrieves the PivotTable named PivotTable1 (noting that timeline creation is not supported), and saves the workbook as SalesReport.pdf.
    class TimelineToPdfExample
    {
        static void Main()
        {
            const string inputFile = "SalesData.xlsx";
            const string outputFile = "SalesReport.pdf";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file \"{inputFile}\" was not found.");
                return;
            }

            try
            {
                // Load the existing Excel workbook
                Workbook workbook = new Workbook(inputFile);

                // Access the first worksheet (or specify by name)
                Worksheet sheet = workbook.Worksheets[0]; // or workbook.Worksheets["Sheet1"]

                // Retrieve the PivotTable named "PivotTable1"
                PivotTable pivotTable = sheet.PivotTables["PivotTable1"];
                if (pivotTable == null)
                {
                    Console.WriteLine("Error: PivotTable \"PivotTable1\" not found on the worksheet.");
                    return;
                }

                // NOTE: Timeline feature may not be available in the current Aspose.Cells version.
                // If needed, timeline creation can be added when the API is supported.

                // Export the workbook (or the specific worksheet) to PDF
                workbook.Save(outputFile, SaveFormat.Pdf);
                Console.WriteLine($"PDF report generated successfully: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
