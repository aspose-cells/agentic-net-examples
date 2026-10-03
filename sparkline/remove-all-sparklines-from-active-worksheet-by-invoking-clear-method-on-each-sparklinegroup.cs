// Title: Remove all sparklines from the active worksheet of an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, clears every SparklineGroup on the active worksheet, and saves the modified workbook to a new path. | Show a C# example that verifies the input file exists, accesses Workbook.Worksheets[Workbook.Worksheets.ActiveSheetIndex], calls SparklineGroups.Clear(), and includes basic exception handling.
// Common Searches: C# Aspose.Cells delete sparkline groups on the current sheet | how to clear sparklines programmatically with Aspose.Cells .NET | Aspose.Cells SparklineGroups.Clear usage example | remove sparklines from an Excel workbook using C# Aspose.Cells
// Tags: SparklineGroups.Clear Aspose.Cells | clear sparkline groups C# | Aspose.Cells delete all sparklines .NET | remove sparklines from active worksheet | Workbook.Save after sparkline removal

using System;
using System.IO;
using Aspose.Cells;

// The program loads "input.xlsx" with Aspose.Cells, checks that the file exists, clears all SparklineGroup objects on the active worksheet via SparklineGroups.Clear(), saves the updated workbook as "output.xlsx", and handles any runtime exceptions.
class RemoveSparklines
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the active worksheet
            Worksheet activeSheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

            // Remove all sparkline groups from the active worksheet
            activeSheet.SparklineGroups.Clear();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Sparklines removed and workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
