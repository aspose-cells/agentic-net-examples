// Title: Remove all sparklines from a worksheet with Aspose.Cells C# by clearing each SparklineGroup
// AI Prompts: Generate C# code that loads a workbook, iterates through every SparklineGroup on the active sheet, and calls Clear on each group's Sparklines collection. | Show how to delete every sparkline in an Aspose.Cells worksheet while preserving other data, then save the workbook to a new file. | Provide a snippet that adds a sample sparkline group, clears all sparklines, and exports the result as an XLSX file.
// Common Searches: Aspose.Cells C# remove all sparklines from a worksheet | clear sparklines in each SparklineGroup using Aspose.Cells API | C# code to delete sparkline groups in an Excel file with Aspose.Cells | how to clear Sparkline collection in Aspose.Cells without losing worksheet data
// Tags: SparklineGroup.Sparklines.Clear C# | Aspose.Cells remove sparkline groups | clear sparklines collection Aspose.Cells | Excel workbook sparkline deletion C# | Aspose.Cells workbook save after sparkline clear

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates (or loads) a workbook, optionally adds a sparkline group, then iterates over all SparklineGroup objects on the first worksheet and calls the Clear method on each group's Sparklines collection. Finally, it saves the workbook as RemovedAllSparklines.xlsx, leaving the worksheet without any sparklines.
    public class RemoveAllSparklines
    {
        // Entry point for the console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // create

            // Access the first worksheet (active worksheet)
            Worksheet sheet = workbook.Worksheets[0];

            // Example: add a sparkline group so we have something to clear
            // This part can be omitted if the worksheet already contains sparklines
            CellArea sparklineLocation = new CellArea
            {
                StartColumn = 4,
                EndColumn = 4,
                StartRow = 0,
                EndRow = 0
            };
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, sparklineLocation);
            SparklineGroup exampleGroup = sheet.SparklineGroups[groupIndex];
            exampleGroup.Sparklines.Add($"{sheet.Name}!A1:D1", 0, 4);

            // Iterate through each SparklineGroup in the worksheet
            // and clear its SparklineCollection using the Clear method
            foreach (SparklineGroup group in sheet.SparklineGroups)
            {
                // Clear all sparklines within the current group
                group.Sparklines.Clear();
            }

            // Save the workbook with the sparklines removed
            workbook.Save("RemovedAllSparklines.xlsx"); // save
        }
    }
}
