// Title: Set a custom blue color for negative points in a line sparkline located in cell J6 using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# how to color negative points of a line sparkline blue | example of applying a custom CellsColor to negative values in a sparkline in cell J6 | change negative sparkline point color with Aspose.Cells .NET API | C# code to set negative points color for a sparkline group in Aspose.Cells
// Tags: line sparkline negative points custom color | Aspose.Cells sparkline group negative points CellsColor | sparkline formatting cell J6 Aspose.Cells | C# Aspose.Cells set negative points color | custom blue color for sparkline negative values

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing; // Required for CellsColor

namespace SparklineNegativePointsDemo
{
    // The example creates a workbook, fills column A with mixed values, adds a line sparkline to cell J6 referencing A1:A5, creates a blue CellsColor, assigns it to the sparkline group's NegativePointsColor property, and saves the file as SparklineNegativeBlue.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data with both positive and negative values in column A (A1:A5)
                double[] sampleData = { 12, -7, 15, -3, 20 };
                for (int i = 0; i < sampleData.Length; i++)
                {
                    worksheet.Cells[i, 0].PutValue(sampleData[i]); // Column 0 = "A"
                }

                // Define the data range for the sparkline and the destination cell (J6)
                string dataRange = "A1:A5";
                CellArea sparklineLocation = CellArea.CreateCellArea("J6", "J6");

                // Add a line sparkline group
                int groupIndex = worksheet.SparklineGroups.Add(
                    0,                 // Line sparkline type code
                    dataRange,         // Source data range
                    false,             // Do not plot empty cells as zero
                    sparklineLocation  // Destination cell area
                );

                // Retrieve the created SparklineGroup
                SparklineGroup sparklineGroup = worksheet.SparklineGroups[groupIndex];

                // Create a CellsColor object for the negative points (custom blue)
                CellsColor negativeColor = workbook.CreateCellsColor();
                negativeColor.Color = Color.Blue;

                // Apply the custom color to negative points
                sparklineGroup.NegativePointsColor = negativeColor;

                // Save the workbook
                string outputPath = "SparklineNegativeBlue.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
                Console.WriteLine("Sparkline in cell J6 now displays negative values in blue.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
