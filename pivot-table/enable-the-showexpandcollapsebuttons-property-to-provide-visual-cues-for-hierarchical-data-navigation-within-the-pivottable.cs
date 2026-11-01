// Title: How to activate ShowExpandCollapseButtons for a PivotTable in an existing Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that loads a workbook, finds the first pivot table, and sets its ShowExpandCollapseButtons flag to true before saving. | Write a C# routine that uses reflection to apply ShowExpandCollapseButtons via the Options object when present, and falls back to the direct property if not. | Create a C# loop that iterates over all pivot tables in a worksheet and enables hierarchical expand/collapse buttons for each using Aspose.Cells.
// Common Searches: aspnet aspose.cells enable expand collapse buttons on pivot table programmatically | c# set ShowExpandCollapseButtons property for pivot tables in existing Excel file | how to show hierarchical expand/collapse icons in Excel pivot table using Aspose.Cells | fallback to Options.ShowExpandCollapseButtons with reflection Aspose.Cells C#
// Tags: Aspose.Cells pivot table expand collapse flag | C# enable hierarchical navigation in pivot tables | reflection fallback for Options property Aspose.Cells | modify existing workbook pivot table settings | Excel pivot table show expand collapse buttons

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, checks for pivot tables on the first worksheet, and uses reflection to enable the ShowExpandCollapseButtons feature—first trying the Options object if it exists, otherwise setting the property directly—then saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the existing workbook
                var workbook = new Workbook(inputPath);

                // Access the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Get the collection of pivot tables in the worksheet
                var pivotTables = worksheet.PivotTables;

                // Ensure there is at least one pivot table
                if (pivotTables.Count > 0)
                {
                    // Access the first pivot table
                    var pivotTable = pivotTables[0];

                    // Enable the expand/collapse buttons (property available in recent API versions)
                    // If the Options property is unavailable, use the direct property.
                    try
                    {
                        // Attempt to use Options if present (for newer versions)
                        var optionsProperty = pivotTable.GetType().GetProperty("Options");
                        if (optionsProperty != null)
                        {
                            var options = optionsProperty.GetValue(pivotTable);
                            var showProp = options.GetType().GetProperty("ShowExpandCollapseButtons");
                            if (showProp != null)
                            {
                                showProp.SetValue(options, true);
                            }
                        }
                        else
                        {
                            // Fallback to direct property for older versions
                            var showProp = pivotTable.GetType().GetProperty("ShowExpandCollapseButtons");
                            if (showProp != null && showProp.CanWrite)
                            {
                                showProp.SetValue(pivotTable, true);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to set expand/collapse buttons: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                }

                // Ensure the output directory exists
                var outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
