// Title: Programmatically enable hidden cells display for all sparkline groups in an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to show hidden cells in sparkline groups with Aspose.Cells C# | set DisplayHidden property for sparkline groups using Aspose.Cells .NET | Aspose.Cells reflection example to modify sparkline settings | enable hidden data points in Excel sparklines via C# code | update sparkline group properties in an existing workbook with Aspose.Cells
// Tags: Aspose.Cells set sparkline DisplayHidden property | C# reflection modify Excel sparkline groups | enable hidden cells in Excel sparklines programmatically | load and save workbook with updated sparkline settings | Aspose.Cells sparkline group configuration .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, uses reflection to iterate over each SparklineGroup in the first worksheet, sets the DisplayHidden property to true for all groups, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Attempt to enable display of hidden cells for any sparkline groups using reflection
                try
                {
                    var sparklineGroupsProp = worksheet.GetType().GetProperty("SparklineGroups");
                    if (sparklineGroupsProp != null)
                    {
                        var sparklineGroups = sparklineGroupsProp.GetValue(worksheet, null);
                        if (sparklineGroups != null)
                        {
                            var countProp = sparklineGroups.GetType().GetProperty("Count");
                            var itemProp = sparklineGroups.GetType().GetProperty("Item"); // indexer
                            if (countProp != null && itemProp != null)
                            {
                                int count = (int)countProp.GetValue(sparklineGroups);
                                for (int i = 0; i < count; i++)
                                {
                                    var group = itemProp.GetValue(sparklineGroups, new object[] { i });
                                    var displayHiddenProp = group?.GetType().GetProperty("DisplayHidden");
                                    if (displayHiddenProp != null && displayHiddenProp.CanWrite)
                                    {
                                        displayHiddenProp.SetValue(group, true);
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception reflEx)
                {
                    // If reflection fails, log but continue processing
                    Console.WriteLine($"Warning: Unable to modify sparkline groups. {reflEx.Message}");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display an error message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
