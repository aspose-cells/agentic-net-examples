// Title: C# – Export Named Ranges and Their RefersTo Formulas to a Summary Sheet with Aspose.Cells
// Description: This example loads an existing workbook (or creates a new one), adds a "Summary" worksheet, writes column headers, iterates the workbook's NameCollection to capture each defined name and its RefersTo formula, and saves the result as a new file. It also ensures the output folder exists and handles missing source files gracefully.
// Keywords: Aspose.Cells | C# | .NET | named ranges | RefersTo formula | export to summary sheet | list defined names | Excel automation | GitHub sample | code example
// Common Searches: Aspose.Cells list all named ranges with formulas C# | export named ranges to a new worksheet using .NET | how to get RefersTo property of defined names in Excel | create summary sheet for named ranges Aspose.Cells | C# code sample for enumerating workbook names collection
// Developer Intent: Create a worksheet that enumerates every named range together with its RefersTo expression.
// Use Cases: Generate a documentation tab that audits all named ranges before publishing the workbook. | Provide end‑users a quick reference sheet showing each range’s purpose and formula. | Export named‑range definitions for version‑control diff or automated testing.
// AI Prompts: Write C# code using Aspose.Cells that adds a "Summary" sheet listing each defined name and its RefersTo formula, with fallback when the source file is missing. | Extend the sample to also output the scope (worksheet or workbook) of each named range in additional columns. | Create a reusable method that receives a Workbook object and returns a DataTable containing Name, RefersTo, and Scope for all named ranges.

using System;
using System.IO;
using Aspose.Cells;

// This example loads an existing workbook (or creates a new one), adds a "Summary" worksheet, writes column headers, iterates the workbook's NameCollection to capture each defined name and its RefersTo formula, and saves the result as a new file. It also ensures the output folder exists and handles missing source files gracefully.
class ExportNamedRangesSummary
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load workbook if the input file exists; otherwise create a new workbook
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Add a new worksheet to hold the summary
            Worksheet summarySheet = workbook.Worksheets.Add("Summary");

            // Write header titles
            summarySheet.Cells["A1"].PutValue("Named Range");
            summarySheet.Cells["B1"].PutValue("Refers To Formula");

            // Retrieve all defined names (named ranges) in the workbook
            NameCollection namedRanges = workbook.Worksheets.Names;

            // Populate the summary sheet if there are any named ranges
            if (namedRanges != null && namedRanges.Count > 0)
            {
                int rowIndex = 2; // Excel rows are 1‑based; row 1 holds headers
                foreach (Name nr in namedRanges)
                {
                    // 'Text' property holds the name of the defined range
                    summarySheet.Cells[rowIndex, 0].PutValue(nr.Text);
                    summarySheet.Cells[rowIndex, 1].PutValue(nr.RefersTo);
                    rowIndex++;
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
