// Title: Export Worksheet Paper Dimensions to CSV with Aspose.Cells for .NET (C#)
// Description: A complete C# example that loads a workbook, reads each worksheet's PageSetup (paper width, height, size, and orientation), writes the data to a new sheet, and saves it as a CSV file for downstream analysis or reporting.
// Keywords: Aspose.Cells | C# | .NET | export to CSV | worksheet paper dimensions | PageSetup properties | PaperWidth | PaperHeight | PaperSize | Orientation | code example | GitHub sample | document formatting report
// Common Searches: Aspose.Cells export worksheet page setup to CSV | C# get paper size of each worksheet | save worksheet dimensions as CSV file | how to read worksheet orientation with Aspose.Cells | extract page setup data from Excel using .NET
// Developer Intent: Retrieve the paper size and orientation of every worksheet in an Excel file and output the information to a CSV document.
// Use Cases: Generate a printable‑settings audit report for all sheets before batch printing. | Feed worksheet dimension data into an analytics pipeline that evaluates formatting consistency across workbooks. | Create a CSV inventory of page‑setup configurations for compliance checks in a document management system.
// AI Prompts: Write C# code using Aspose.Cells that extracts PaperWidth, PaperHeight, PaperSize, and Orientation from each worksheet and saves the results to a CSV file. | Extend the sample to also include left, right, top, and bottom margin values in the CSV output. | Explain how to add a column for each worksheet's default print area while keeping the CSV format.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // A complete C# example that loads a workbook, reads each worksheet's PageSetup (paper width, height, size, and orientation), writes the data to a new sheet, and saves it as a CSV file for downstream analysis or reporting.
    public class ExportWorksheetPaperDimensions
    {
        public static void Run()
        {
            try
            {
                string sourcePath = "source.xlsx";

                // Verify that the source file exists before loading
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file '{sourcePath}' not found.");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Create a new workbook to hold the dimensions data
                Workbook dimensionsWorkbook = new Workbook();
                Worksheet sheet = dimensionsWorkbook.Worksheets[0];
                sheet.Name = "PaperDimensions";

                // Write header row
                sheet.Cells["A1"].PutValue("Worksheet");
                sheet.Cells["B1"].PutValue("PaperWidth (inches)");
                sheet.Cells["C1"].PutValue("PaperHeight (inches)");
                sheet.Cells["D1"].PutValue("PaperSize");
                sheet.Cells["E1"].PutValue("Orientation");

                // Iterate through each worksheet in the source workbook
                for (int i = 0; i < sourceWorkbook.Worksheets.Count; i++)
                {
                    Worksheet ws = sourceWorkbook.Worksheets[i];
                    PageSetup ps = ws.PageSetup;

                    // Row index in the dimensions sheet (starting from row 2)
                    int row = i + 1;

                    // Fill data
                    sheet.Cells[row, 0].PutValue(ws.Name);
                    sheet.Cells[row, 1].PutValue(ps.PaperWidth);
                    sheet.Cells[row, 2].PutValue(ps.PaperHeight);
                    sheet.Cells[row, 3].PutValue(ps.PaperSize.ToString());
                    sheet.Cells[row, 4].PutValue(ps.Orientation.ToString());
                }

                // Save the dimensions workbook as CSV
                string outputPath = "WorksheetPaperDimensions.csv";
                dimensionsWorkbook.Save(outputPath, SaveFormat.Csv);
                Console.WriteLine($"Dimensions saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportWorksheetPaperDimensions.Run();
        }
    }
}
