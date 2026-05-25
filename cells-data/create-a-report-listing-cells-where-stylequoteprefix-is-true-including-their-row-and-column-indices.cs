using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixReport
{
    public class ReportGenerator
    {
        public static void Run()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output_with_quoteprefix_report.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the source workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (source data)
                Worksheet sourceSheet = workbook.Worksheets[0];
                Cells sourceCells = sourceSheet.Cells;

                // Create a new worksheet to hold the report
                Worksheet reportSheet = workbook.Worksheets.Add("QuotePrefixReport");
                Cells reportCells = reportSheet.Cells;

                // Write header row in the report sheet
                reportCells[0, 0].PutValue("Row Index");
                reportCells[0, 1].PutValue("Column Index");
                reportCells[0, 2].PutValue("Cell Address");
                reportCells[0, 3].PutValue("Cell Value");

                int reportRow = 1; // Start writing data from the second row

                // Determine the used range of the source sheet
                int maxRow = sourceCells.MaxDataRow;
                int maxCol = sourceCells.MaxDataColumn;

                // Iterate through each cell in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sourceCells[row, col];

                        // Skip null cells or cells without a value
                        if (cell == null || cell.Value == null)
                            continue;

                        // Retrieve the cell's style and check the QuotePrefix property
                        Style style = cell.GetStyle();
                        if (style.QuotePrefix)
                        {
                            // Record the cell information in the report sheet
                            reportCells[reportRow, 0].PutValue(row);               // Row index (0‑based)
                            reportCells[reportRow, 1].PutValue(col);               // Column index (0‑based)
                            reportCells[reportRow, 2].PutValue(cell.Name);        // Excel address (e.g., "B10")
                            reportCells[reportRow, 3].PutValue(cell.Value);       // Actual cell value
                            reportRow++;
                        }
                    }
                }

                // Save the workbook with the added report sheet
                workbook.Save(outputPath);
                Console.WriteLine($"Report generated and saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ReportGenerator.Run();
        }
    }
}