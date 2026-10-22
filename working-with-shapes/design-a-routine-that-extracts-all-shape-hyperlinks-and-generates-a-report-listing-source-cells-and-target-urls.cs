// Title: Generate a report of Excel shape hyperlinks with their source cells using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, scans every worksheet for shapes, extracts each shape's Hyperlink.Address and the cell coordinates where the shape is anchored, and writes the worksheet name, cell address, and URL to a new worksheet in a separate workbook. | Create a routine in C# using Aspose.Cells that iterates through all shapes in a workbook, checks for non‑empty Hyperlink objects, captures the UpperLeftRow/UpperLeftColumn to form the cell reference, and saves a two‑column report (SourceCell, TargetUrl) to an output Excel file.
// Common Searches: Aspose.Cells C# extract hyperlink from shape and get its cell address | How to list all shape hyperlinks in an Excel file using .NET | Create a hyperlink report for shapes in Excel with Aspose.Cells | Get source cell of a shape that contains a hyperlink in Aspose.Cells | Export shape hyperlink data to a new workbook using C#
// Tags: Aspose.Cells extract shape links | C# get shape anchor cell address | create Excel hyperlink report with Aspose.Cells | iterate worksheets and shapes using Aspose.Cells | export hyperlink data to new workbook C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeHyperlinkExtractor
{
    // The program loads an input Excel file, iterates through each worksheet and its shapes, identifies shapes that contain hyperlinks, records the worksheet name and cell address where each shape is anchored along with the hyperlink URL, and writes these pairs to a new workbook named "hyperlink_report.xlsx".
    class Program
    {
        static void Main(string[] args)
        {
            // Input workbook containing shapes with hyperlinks
            string inputFilePath = "input.xlsx";

            // Output workbook that will hold the report
            string reportFilePath = "hyperlink_report.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputFilePath))
                {
                    Console.WriteLine($"Input file not found: {inputFilePath}");
                    return;
                }

                // Load the source workbook (lifecycle rule: load)
                Workbook sourceWorkbook = new Workbook(inputFilePath);

                // Create a new workbook for the report (lifecycle rule: create)
                Workbook reportWorkbook = new Workbook();

                // Prepare the first worksheet of the report
                Worksheet reportSheet = reportWorkbook.Worksheets[0];
                reportSheet.Name = "Hyperlink Report";

                // Write header row
                reportSheet.Cells[0, 0].PutValue("Source Cell");
                reportSheet.Cells[0, 1].PutValue("Target URL");

                int reportRowIndex = 1; // start after header

                // Iterate through all worksheets in the source workbook
                foreach (Worksheet ws in sourceWorkbook.Worksheets)
                {
                    // Iterate through all shapes in the current worksheet
                    foreach (Shape shape in ws.Shapes)
                    {
                        // Check if the shape contains a hyperlink
                        if (shape.Hyperlink != null && !string.IsNullOrEmpty(shape.Hyperlink.Address))
                        {
                            // Determine the cell address where the shape is anchored
                            int row = shape.UpperLeftRow;          // zero‑based row index
                            int column = shape.UpperLeftColumn;   // zero‑based column index
                            string cellAddress = CellsHelper.CellIndexToName(row, column);

                            // Combine worksheet name and cell address for clarity
                            string sourceCell = $"{ws.Name}!{cellAddress}";

                            // Retrieve the target URL from the hyperlink
                            string targetUrl = shape.Hyperlink.Address;

                            // Write the information into the report worksheet
                            reportSheet.Cells[reportRowIndex, 0].PutValue(sourceCell);
                            reportSheet.Cells[reportRowIndex, 1].PutValue(targetUrl);
                            reportRowIndex++;
                        }
                    }
                }

                // Save the report workbook (lifecycle rule: save)
                reportWorkbook.Save(reportFilePath);
                Console.WriteLine($"Hyperlink report saved to: {reportFilePath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
