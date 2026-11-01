// Title: Identify and log differences in paper size, orientation, and margins between two Excel worksheets using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that loads a workbook, reads the PageSetup of two specified worksheets, and records any mismatches in PaperSize, Orientation, and all margin values. | Create a reusable method that compares two double margin values with a configurable tolerance and adds a descriptive message to a differences collection.
// Common Searches: how to compare page setup properties of two sheets with Aspose.Cells C# | C# program to find differences in paper size and margins between Excel worksheets | log worksheet layout discrepancies using Aspose.Cells .NET API | detect orientation mismatch between Excel sheets in C# | compare top and bottom margins of two worksheets with tolerance Aspose.Cells
// Tags: worksheet page setup comparison Aspose.Cells | paper size mismatch detection C# | margin tolerance evaluation Excel | orientation difference logging .NET | compare sheet layout properties Aspose

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel workbook, accesses two worksheets, and compares their PageSetup settings—including PaperSize, Orientation, and all margin values—using a small tolerance for margins. Any differences are collected, printed to the console, and the workbook can be saved.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the file
            var workbook = new Workbook(inputPath);

            // Retrieve worksheets by name; ensure they exist
            var sheet1 = workbook.Worksheets["Sheet1"];
            var sheet2 = workbook.Worksheets["Sheet2"];

            if (sheet1 == null || sheet2 == null)
            {
                Console.WriteLine("One or both worksheets (Sheet1, Sheet2) were not found in the workbook.");
                return;
            }

            // List to hold any differences found
            var differences = new List<string>();

            // Compare PaperSize
            if (sheet1.PageSetup.PaperSize != sheet2.PageSetup.PaperSize)
            {
                differences.Add($"PaperSize differs: Sheet1={sheet1.PageSetup.PaperSize}, Sheet2={sheet2.PageSetup.PaperSize}");
            }

            // Compare Orientation
            if (sheet1.PageSetup.Orientation != sheet2.PageSetup.Orientation)
            {
                differences.Add($"Orientation differs: Sheet1={sheet1.PageSetup.Orientation}, Sheet2={sheet2.PageSetup.Orientation}");
            }

            // Compare margins (values are in points)
            CompareMargin("TopMargin", sheet1.PageSetup.TopMargin, sheet2.PageSetup.TopMargin, differences);
            CompareMargin("BottomMargin", sheet1.PageSetup.BottomMargin, sheet2.PageSetup.BottomMargin, differences);
            CompareMargin("LeftMargin", sheet1.PageSetup.LeftMargin, sheet2.PageSetup.LeftMargin, differences);
            CompareMargin("RightMargin", sheet1.PageSetup.RightMargin, sheet2.PageSetup.RightMargin, differences);

            // Output the results
            if (differences.Count == 0)
            {
                Console.WriteLine("No differences in paper dimensions were found.");
            }
            else
            {
                Console.WriteLine("Differences found:");
                foreach (var diff in differences)
                {
                    Console.WriteLine(diff);
                }
            }

            // Save the workbook (optional, if modifications were made)
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to compare margin values with a tolerance
    static void CompareMargin(string marginName, double value1, double value2, List<string> diffs)
    {
        const double tolerance = 0.01; // points tolerance
        if (Math.Abs(value1 - value2) > tolerance)
        {
            diffs.Add($"{marginName} differs: Sheet1={value1}, Sheet2={value2}");
        }
    }
}
