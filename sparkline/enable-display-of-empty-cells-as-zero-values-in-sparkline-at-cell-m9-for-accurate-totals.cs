// Title: How to Configure a Sparkline to Show Empty Cells as Zero Values Using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set sparkline empty points to zero | display blank cells as zero in a line sparkline with Aspose.Cells .NET | SparklineGroup.EmptyPointsDisplay Zero example in C# | show zero for missing data in Excel sparkline using Aspose.Cells | configure sparkline to treat empty cells as zero values in .xlsx with Aspose.Cells
// Tags: set sparkline empty points to zero | sparkline empty points display property | Aspose.Cells line sparkline configuration | C# sparkline zero values handling | Excel sparkline blank cells zero .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Load the workbook, define the data range, add a line sparkline to cell M9, and set SparklineGroup.EmptyPointsDisplay = EmptyPointsDisplay.Zero so that any empty cells in the source range are rendered as zero values, ensuring correct total calculations.
class SparklineExample
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Define the data range for the sparkline (e.g., cells B2:B10)
            // Rows and columns are zero‑based (B2 = row 1, column 1)
            CellArea dataRange = new CellArea
            {
                StartRow = 1,
                StartColumn = 1,
                EndRow = 9,
                EndColumn = 1
            };

            // Add a sparkline group to the worksheet.
            // The sparkline will be placed in cell M9.
            // The third argument (false) indicates a horizontal sparkline.
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "M9", false, dataRange);
            SparklineGroup sparklineGroup = sheet.SparklineGroups[groupIndex];

            // Optional: customize other sparkline appearance settings
            // sparklineGroup.LineWeight = 0.5;
            // sparklineGroup.Color = System.Drawing.Color.Blue;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
