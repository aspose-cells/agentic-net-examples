// Title: Create an Excel slicer audit worksheet that lists slicer caption, size, and linked table column using Aspose.Cells for .NET
// Common Searches: how to list all slicer captions, widths and heights in an Excel file with Aspose.Cells C# | Aspose.Cells C# generate report of slicer dimensions and linked table column | extract slicer properties to a new worksheet using Aspose.Cells for .NET | C# code to audit slicers in a workbook and save details to another sheet | retrieve slicer TableName and ColumnName with Aspose.Cells API
// Tags: Aspose.Cells slicer enumeration | C# slicer dimension extraction | Excel slicer linked table metadata | Aspose.Cells generate audit worksheet | C# workbook slicer report

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The sample loads an existing workbook, adds a worksheet named "SlicerReport", writes headers, then iterates through every worksheet's slicer collection. For each slicer it captures the Caption, Width, Height, TableName and ColumnName, writes these values to the report sheet, and finally saves the workbook as a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Add a new worksheet for the slicer audit report
            int reportIndex = workbook.Worksheets.Add();
            Worksheet reportSheet = workbook.Worksheets[reportIndex];
            reportSheet.Name = "SlicerReport";

            // Write header row
            reportSheet.Cells["A1"].PutValue("Slicer Title");
            reportSheet.Cells["B1"].PutValue("Width");
            reportSheet.Cells["C1"].PutValue("Height");
            reportSheet.Cells["D1"].PutValue("Linked Table.Column");

            int currentRow = 1; // zero‑based index (row 2 in Excel)

            // Iterate through all worksheets and their slicers
            foreach (Worksheet ws in workbook.Worksheets)
            {
                try
                {
                    // Use dynamic to avoid compile‑time dependency on Slicer types
                    dynamic slicerCollection = ws.Slicers;

                    foreach (dynamic slicer in slicerCollection)
                    {
                        // Title (caption) of the slicer
                        string title = slicer.Caption ?? string.Empty;

                        // Size of the slicer
                        double width = slicer.Width;
                        double height = slicer.Height;

                        // Linked table and column (e.g., Table1.ColumnA)
                        string linkedInfo = $"{slicer.TableName}.{slicer.ColumnName}";

                        // Populate the report row
                        reportSheet.Cells[currentRow, 0].PutValue(title);
                        reportSheet.Cells[currentRow, 1].PutValue(width);
                        reportSheet.Cells[currentRow, 2].PutValue(height);
                        reportSheet.Cells[currentRow, 3].PutValue(linkedInfo);

                        currentRow++;
                    }
                }
                catch (Exception ex)
                {
                    // If a worksheet has no slicers or an unexpected error occurs, log and continue
                    Console.WriteLine($"Warning: Could not process slicers in worksheet '{ws.Name}'. {ex.Message}");
                }
            }

            // Save the workbook with the new report sheet
            workbook.Save(outputPath);
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
