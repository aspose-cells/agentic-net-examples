using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklinesDemo
{
    class Program
    {
        static void Main()
        {
            // Define the output file path (XLSX format)
            string outputPath = @"C:\Temp\SparklinesOutput.xlsx";

            try
            {
                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Create a new workbook
                using (Workbook workbook = new Workbook())
                {
                    // Access the first worksheet
                    Worksheet sheet = workbook.Worksheets[0];

                    // Populate sample data for the sparkline (row 1, columns A to D)
                    sheet.Cells["A1"].PutValue(5);
                    sheet.Cells["B1"].PutValue(12);
                    sheet.Cells["C1"].PutValue(8);
                    sheet.Cells["D1"].PutValue(15);

                    // Define the cell area where the sparkline will be placed (E1)
                    CellArea sparklineLocation = new CellArea
                    {
                        StartRow = 0,
                        EndRow = 0,
                        StartColumn = 4, // Column E (0‑based index)
                        EndColumn = 4
                    };

                    // Add a sparkline group of type Line using the data range A1:D1
                    int groupIndex = sheet.SparklineGroups.Add(
                        SparklineType.Line,
                        sheet.Name + "!A1:D1",   // data range
                        false,                   // show markers
                        sparklineLocation);      // location of the sparkline

                    // Retrieve the created group (optional, for further customization)
                    SparklineGroup group = sheet.SparklineGroups[groupIndex];

                    // Example: set the sparkline style (optional)
                    group.LineWeight = 0.75;
                    // Note: SparklineGroup does not expose a direct Color property in this API version.
                    // Additional styling can be applied via other available properties if needed.

                    // Save the workbook to the specified path as XLSX
                    workbook.Save(outputPath);
                }

                Console.WriteLine($"Workbook with sparklines saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}