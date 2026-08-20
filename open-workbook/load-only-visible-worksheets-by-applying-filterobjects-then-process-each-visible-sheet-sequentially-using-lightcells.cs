// Title: C# – Load Only Visible Worksheets and Process Cells with LightCells using Aspose.Cells
// Description: Demonstrates how to create a custom LoadFilter that loads data only from visible worksheets, retrieve their indexes, and sequentially process each sheet's cells with the LightCells API. This approach minimizes memory usage by skipping hidden sheets while providing full cell access for reporting or transformation tasks.
// Keywords: Aspose.Cells C# load visible worksheets | custom LoadFilter hidden sheets | LightCells process visible sheets | skip hidden worksheets Aspose.Cells | memory‑efficient workbook loading | iterate cells visible worksheets | LoadOptions LoadFilter example | GitHub Aspose.Cells sample
// Common Searches: How to load only visible worksheets with Aspose.Cells .NET | Aspose.Cells custom LoadFilter to ignore hidden sheets | Iterate cells of visible worksheets in C# | LightCells API example for visible sheets | Reduce memory usage when opening large Excel files Aspose
// Developer Intent: Load a workbook while excluding hidden worksheets, then loop through every cell of each visible sheet using LightCells for efficient processing.
// Use Cases: Generate reports that include data solely from user‑visible tabs, cutting down on processing time. | Extract or transform data from visible sheets in large workbooks without loading hidden content into memory. | Create automated scripts that scan visible worksheets for specific values or patterns while keeping the footprint low.
// AI Prompts: Write C# code that uses Aspose.Cells LoadFilter to load only visible worksheets and then processes each cell with LightCells. | Show an example of a custom LoadFilter in Aspose.Cells that skips hidden sheets and returns the indexes of visible worksheets. | Provide a method that iterates over cell values of visible worksheets after applying a LoadOptions filter in Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsVisibleSheetsLightCells
{
    // Custom LoadFilter to load only visible worksheets
    // Demonstrates how to create a custom LoadFilter that loads data only from visible worksheets, retrieve their indexes, and sequentially process each sheet's cells with the LightCells API. This approach minimizes memory usage by skipping hidden sheets while providing full cell access for reporting or transformation tasks.
    public class VisibleSheetsLoadFilter : LoadFilter
    {
        public VisibleSheetsLoadFilter() : base(LoadDataFilterOptions.All) { }

        // Adjust loading options based on worksheet visibility
        public override void StartSheet(Worksheet sheet)
        {
            LoadDataFilterOptions = sheet.IsVisible
                ? LoadDataFilterOptions.All
                : LoadDataFilterOptions.Structure;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source workbook.
                string sourceFile = "InputWorkbook.xlsx";

                // Ensure the input file exists to avoid FileNotFoundException.
                if (!File.Exists(sourceFile))
                {
                    Console.WriteLine($"Error: File \"{sourceFile}\" not found.");
                    return;
                }

                // Configure load options with the custom filter.
                var loadOptions = new LoadOptions
                {
                    LoadFilter = new VisibleSheetsLoadFilter()
                };

                // Load the workbook using the configured options.
                // Only visible worksheets will have their data loaded.
                using (var workbook = new Workbook(sourceFile, loadOptions))
                {
                    // Get indexes of visible sheets.
                    List<int> visibleIndexesList = new List<int>();
                    for (int i = 0; i < workbook.Worksheets.Count; i++)
                    {
                        if (workbook.Worksheets[i].IsVisible)
                            visibleIndexesList.Add(i);
                    }
                    int[] visibleIndexes = visibleIndexesList.ToArray();

                    // List visible sheets.
                    Console.WriteLine("\nVisible sheets in the workbook:");
                    foreach (int index in visibleIndexes)
                    {
                        Worksheet ws = workbook.Worksheets[index];
                        Console.WriteLine($"- {ws.Name}");
                    }

                    // Process cells of visible worksheets.
                    Console.WriteLine("\nProcessing cells of visible sheets:");
                    foreach (int index in visibleIndexes)
                    {
                        Worksheet ws = workbook.Worksheets[index];
                        Console.WriteLine($"Start processing sheet: {ws.Name}");
                        foreach (Cell cell in ws.Cells)
                        {
                            Console.WriteLine($"  Cell {cell.Name}: {cell.StringValue}");
                        }
                    }

                    // Save the workbook if any changes were made (not required for this demo).
                    // workbook.Save("ProcessedWorkbook.xlsx");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
