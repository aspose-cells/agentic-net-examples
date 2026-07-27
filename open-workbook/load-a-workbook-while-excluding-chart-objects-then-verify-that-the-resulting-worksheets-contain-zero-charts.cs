// Title: Load an Excel workbook with Aspose.Cells for .NET, clear all chart objects, and confirm zero charts per worksheet
// Description: C# example that opens a workbook, removes every chart from each worksheet using Aspose.Cells, then iterates through the sheets to verify that the chart count is zero and reports the result.
// Keywords: Aspose.Cells load workbook without charts | clear worksheet charts C# | remove chart objects Aspose.Cells | check chart count Aspose.Cells | verify no charts Excel .NET | Excel chart removal programmatically
// Common Searches: how to delete all charts from an Excel file using Aspose.Cells | Aspose.Cells remove charts from workbook C# | verify that a workbook has no charts after loading | C# code to clear charts in each worksheet | Aspose.Cells chart count per worksheet
// Developer Intent: Open a workbook, strip all chart objects, and ensure each worksheet contains zero charts.
// Use Cases: Sanitize user‑uploaded Excel files by stripping charts for security or compliance. | Prepare a template workbook for data‑only processing where charts are unnecessary. | Convert workbooks to chart‑free formats (e.g., CSV) by removing visual objects programmatically.
// AI Prompts: Write a reusable C# method that takes a file path, loads the workbook with Aspose.Cells, removes all chart objects, and returns true if no charts remain. | Generate a console‑app snippet that logs the chart count for each worksheet after clearing charts using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// C# example that opens a workbook, removes every chart from each worksheet using Aspose.Cells, then iterates through the sheets to verify that the chart count is zero and reports the result.
class Program
{
    static void Main()
    {
        // Path to the source workbook that contains charts
        string sourcePath = "input_with_charts.xlsx";

        // Ensure the input file exists before attempting to load it
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Error: File not found - \"{sourcePath}\"");
            return;
        }

        try
        {
            // Load the workbook (charts will be loaded initially)
            Workbook workbook = new Workbook(sourcePath);

            // Remove all charts from each worksheet to simulate loading without charts
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.Charts.Clear();
            }

            // Verify that each worksheet now has zero charts
            bool allZero = true;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int chartCount = sheet.Charts.Count;
                Console.WriteLine($"Worksheet \"{sheet.Name}\" chart count: {chartCount}");
                if (chartCount != 0)
                    allZero = false;
            }

            Console.WriteLine(allZero
                ? "All worksheets contain zero charts."
                : "Some worksheets still contain charts.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }
}
