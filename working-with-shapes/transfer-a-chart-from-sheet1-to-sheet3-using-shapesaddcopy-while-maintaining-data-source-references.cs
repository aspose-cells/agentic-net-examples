// Title: Copy a chart from Sheet1 to Sheet3 while preserving its data source using Aspose.Cells Shapes.AddCopy in C#
// Common Searches: how to copy a chart to another worksheet with Aspose.Cells C# preserving data source | using Shapes.AddCopy to move a chart shape between sheets in Aspose.Cells .NET | transfer chart from Sheet1 to Sheet3 without breaking source references Aspose.Cells | duplicate a chart shape while keeping linked data in a workbook using C# | move chart between worksheets in Aspose.Cells without losing data connections
// Tags: Shapes.AddCopy chart duplication | chart shape copy preserving data source | Aspose.Cells transfer chart between worksheets | C# Aspose.Cells chart relocation | worksheet chart shape copy

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example loads a workbook, ensures Sheet1 and Sheet3 exist, retrieves the first chart on Sheet1, obtains its associated shape, and copies that shape to Sheet3 with Shapes.AddCopy, preserving the original data source references. The copied chart is repositioned and the workbook is saved.
class TransferChartExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that contains the source chart
            Workbook workbook = new Workbook(inputPath);

            // Get source and destination worksheets
            Worksheet sourceSheet = workbook.Worksheets["Sheet1"];
            if (sourceSheet == null)
            {
                Console.WriteLine("Source worksheet \"Sheet1\" not found.");
                return;
            }

            Worksheet destSheet = workbook.Worksheets["Sheet3"];
            if (destSheet == null)
            {
                // Create Sheet3 if it does not exist
                int newIndex = workbook.Worksheets.Add();
                destSheet = workbook.Worksheets[newIndex];
                destSheet.Name = "Sheet3";
            }

            // Ensure there is at least one chart on the source sheet
            if (sourceSheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on Sheet1.");
                return;
            }

            // Get the first chart
            Chart sourceChart = sourceSheet.Charts[0];

            // Retrieve the shape that represents the chart (charts are also stored as shapes)
            Shape sourceChartShape = sourceSheet.Shapes[sourceChart.Name];
            if (sourceChartShape == null)
            {
                Console.WriteLine("Corresponding chart shape not found.");
                return;
            }

            // Copy the chart shape to the destination sheet while preserving data source references
            Shape copiedShape = destSheet.Shapes.AddCopy(
                sourceChartShape,
                sourceChartShape.UpperLeftRow,
                sourceChartShape.UpperLeftColumn,
                sourceChartShape.LowerRightRow,
                sourceChartShape.LowerRightColumn);

            // Reposition the copied chart on the destination sheet (5 points = 5/72 inch)
            copiedShape.Left = 5 * 72;
            copiedShape.Top = 5 * 72;

            // Save the workbook with the transferred chart
            workbook.Save(outputPath);
            Console.WriteLine($"Chart transferred successfully. Output saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
