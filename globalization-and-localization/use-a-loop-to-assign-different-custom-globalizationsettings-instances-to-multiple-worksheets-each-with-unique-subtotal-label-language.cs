// Title: Aspose.Cells .NET: Loop to assign language‑specific subtotal labels with custom GlobalizationSettings per worksheet
// Description: Creates a workbook, fills each sheet with sample sales data, and uses a loop to apply a distinct CustomGlobalizationSettings instance that returns a language‑specific total name (e.g., "Total", "Summe", "合計"). The code defines the data range, runs a Subtotal operation grouped by the Region column, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | .NET | GlobalizationSettings | custom subtotal label | multilingual Excel | worksheet loop | subtotal total name | localization | ConsolidationFunction.Sum | Excel automation
// Common Searches: Aspose.Cells set subtotal label per worksheet | custom GlobalizationSettings for different languages | loop create worksheets with localized total names | C# change subtotal total name Aspose.Cells | multilingual subtotal rows Excel Aspose
// Developer Intent: Generate a workbook with several worksheets, each showing the subtotal total label in a different language by assigning a unique CustomGlobalizationSettings instance inside a loop.
// Use Cases: Produce a multilingual sales report where each sheet displays the subtotal label in its native language. | Create regional financial statements that comply with local reporting terminology. | Automate Excel workbook generation for international distribution with language‑specific subtotal headings.
// AI Prompts: Rewrite the code so each worksheet retains its own GlobalizationSettings without being overwritten by later iterations. | Explain how Aspose.Cells selects the total label for subtotals and how to read the current GlobalizationSettings for a worksheet. | Show how to extend CustomGlobalizationSettings to provide custom labels for Average, Count, and other consolidation functions per worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsGlobalizationDemo
{
    // Custom globalization settings that return a language‑specific total name.
    // Creates a workbook, fills each sheet with sample sales data, and uses a loop to apply a distinct CustomGlobalizationSettings instance that returns a language‑specific total name (e.g., "Total", "Summe", "合計"). The code defines the data range, runs a Subtotal operation grouped by the Region column, and saves the file as an Excel workbook.
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        private readonly string _totalName;

        public CustomGlobalizationSettings(string totalName)
        {
            _totalName = totalName;
        }

        // Override the method that provides the total label for Subtotal operations.
        public override string GetTotalName(ConsolidationFunction functionType)
        {
            // For simplicity we ignore the function type and return the language‑specific name.
            return _totalName;
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Sample data that will be used for all worksheets.
                string[,] data = {
                    { "Region", "Sales" },
                    { "North",   "1000" },
                    { "South",   "2000" },
                    { "East",    "3000" },
                    { "West",    "4000" }
                };

                // Define total‑name strings for each worksheet (different languages).
                string[] totalNames = { "Total", "Summe", "合計" };

                // Loop to create worksheets, assign a unique GlobalizationSettings instance,
                // and apply a Subtotal operation that uses the custom total label.
                for (int i = 0; i < totalNames.Length; i++)
                {
                    // Add a new worksheet (the first worksheet already exists, so we reuse it for i==0).
                    Worksheet sheet;
                    if (i == 0)
                    {
                        sheet = workbook.Worksheets[0];
                    }
                    else
                    {
                        // Worksheets.Add() returns the index of the newly added sheet in some versions.
                        int newIndex = workbook.Worksheets.Add();
                        sheet = workbook.Worksheets[newIndex];
                    }

                    // Populate the worksheet with the sample data.
                    for (int r = 0; r < data.GetLength(0); r++)
                    {
                        for (int c = 0; c < data.GetLength(1); c++)
                        {
                            sheet.Cells[r, c].PutValue(data[r, c]);
                        }
                    }

                    // Assign a custom GlobalizationSettings instance for this workbook.
                    // (Aspose.Cells applies globalization settings at the workbook level.)
                    workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings(totalNames[i]);

                    // Define the range that will be subtotaled (A1:B5 in this example).
                    CellArea area = CellArea.CreateCellArea(0, 0, 4, 1);

                    // Apply Subtotal:
                    //   - group by column 0 (Region)
                    //   - use Sum as the consolidation function
                    //   - replace existing data with the subtotal result
                    sheet.Cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);
                }

                // Save the workbook.
                string outputPath = "WorkbookWithMultipleGlobalizations.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
