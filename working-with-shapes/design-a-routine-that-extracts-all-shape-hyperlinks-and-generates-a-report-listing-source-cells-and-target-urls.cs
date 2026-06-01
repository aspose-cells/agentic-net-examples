using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSample
{
    public class ShapeHyperlinkReport
    {
        public static void Run()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output_with_report.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Add a new worksheet to hold the report
                int reportSheetIndex = workbook.Worksheets.Add();
                Worksheet reportSheet = workbook.Worksheets[reportSheetIndex];
                reportSheet.Name = "HyperlinkReport";

                // Write header row
                reportSheet.Cells["A1"].PutValue("Worksheet");
                reportSheet.Cells["B1"].PutValue("Shape Name");
                reportSheet.Cells["C1"].PutValue("Cell");
                reportSheet.Cells["D1"].PutValue("URL");

                int reportRow = 1; // zero‑based index (row 2 in Excel)

                // Iterate through all worksheets except the report sheet itself
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Name == "HyperlinkReport")
                        continue;

                    // Iterate through all shapes on the worksheet
                    foreach (Shape shape in ws.Shapes)
                    {
                        // Get the hyperlink associated with the shape
                        Hyperlink hyperlink = shape.Hyperlink;

                        // If a hyperlink exists and has a valid address, record it
                        if (hyperlink != null && !string.IsNullOrEmpty(hyperlink.Address))
                        {
                            // Determine the top‑left cell where the shape is anchored
                            int startRow = shape.UpperLeftRow;
                            int startColumn = shape.UpperLeftColumn;
                            string cellRef = CellsHelper.CellIndexToName(startRow, startColumn);

                            // Populate the report row
                            reportSheet.Cells[reportRow, 0].PutValue(ws.Name);               // Worksheet name
                            reportSheet.Cells[reportRow, 1].PutValue(shape.Name);           // Shape name
                            reportSheet.Cells[reportRow, 2].PutValue(cellRef);              // Cell reference
                            reportSheet.Cells[reportRow, 3].PutValue(hyperlink.Address);   // Target URL

                            reportRow++;
                        }
                    }
                }

                // Save the workbook with the added report
                workbook.Save(outputPath);
                Console.WriteLine($"Report generated successfully: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    internal class Program
    {
        private static void Main(string[] args)
        {
            ShapeHyperlinkReport.Run();
        }
    }
}