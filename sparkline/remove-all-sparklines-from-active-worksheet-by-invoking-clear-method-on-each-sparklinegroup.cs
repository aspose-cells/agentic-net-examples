// Title: C# – Remove All Sparklines from an Excel Worksheet with Aspose.Cells for .NET
// Description: Demonstrates how to load or create a workbook, access the first worksheet, iterate through each SparklineGroup, clear its Sparklines collection, and save the file without any sparklines using Aspose.Cells for .NET.
// Keywords: Aspose.Cells clear sparklines | C# remove all sparklines | SparklineGroup Clear method | delete Excel sparklines .NET | Aspose.Cells SparklineGroup example
// Common Searches: how to clear sparklines with Aspose.Cells C# | remove all sparkline groups from Excel using .NET | Aspose.Cells delete sparklines programmatically | C# code to strip sparklines from a worksheet
// Developer Intent: Eliminate every sparkline on the active worksheet programmatically.
// Use Cases: Strip visual sparklines from a template before distributing it to clients. | Comply with data‑privacy policies by removing trend indicators that could expose sensitive information. | Reset a sheet's sparklines so new ones can be generated from refreshed data.
// AI Prompts: Write C# code that uses Aspose.Cells to clear all sparklines from the first worksheet and save the workbook. | Explain how to loop through SparklineGroup objects in Aspose.Cells and remove their sparklines without altering other sheet content. | Show the steps to load an existing Excel file, delete every sparkline on sheet 1, and write the result using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to load or create a workbook, access the first worksheet, iterate through each SparklineGroup, clear its Sparklines collection, and save the file without any sparklines using Aspose.Cells for .NET.
class RemoveAllSparklines
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // Replace with new Workbook("input.xlsx") to load

        // Get the first worksheet (active worksheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Iterate through each SparklineGroup in the worksheet
        foreach (SparklineGroup group in sheet.SparklineGroups)
        {
            // Clear all sparklines within the current group
            group.Sparklines.Clear();
        }

        // Save the workbook after removing sparklines
        workbook.Save("output_without_sparklines.xlsx");
    }
}
