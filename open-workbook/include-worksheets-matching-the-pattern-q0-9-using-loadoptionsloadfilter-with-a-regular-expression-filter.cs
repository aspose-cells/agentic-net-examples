// Title: Load only worksheets named Q followed by numbers using Aspose.Cells LoadOptions.LoadFilter with a regular expression in C#
// AI Prompts: Load an Excel file with Aspose.Cells using LoadOptions.LoadFilter and the regex ^Q[0-9]+$ so that only matching worksheets are loaded, then save the workbook. | Replace the manual sheet‑removal loop with a LoadFilter that selects worksheets whose names start with "Q" and contain only digits. | Create a helper method that accepts a file path and a regex pattern, loads only the matching sheets via LoadOptions.LoadFilter, and returns the filtered Workbook object.
// Common Searches: Aspose.Cells C# load workbook with regex to include specific worksheets | How to use LoadOptions.LoadFilter for sheet name pattern matching in Aspose.Cells | Filter Excel worksheets by name using a regular expression when opening with Aspose.Cells | Load only sheets named Q1, Q2, etc., using Aspose.Cells LoadFilter | Example of regex sheet selection with Aspose.Cells LoadOptions in C#
// Tags: loadfilter regex aspose.cells | worksheet name regex filter aspose.cells c# | select specific sheets during workbook load aspose.cells | exclude nonmatching worksheets aspose.cells | regex based sheet loading aspose.cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The program demonstrates how to open an Excel workbook with Aspose.Cells, apply LoadOptions.LoadFilter using a regular expression (^Q[0-9]+$) to load only worksheets whose names match the pattern, and then save the filtered workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Define the path to the source workbook.
            string sourcePath = "input.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException.
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook without any filter.
            Workbook workbook = new Workbook(sourcePath);

            // Regular expression to match worksheet names like "Q1", "Q23", etc.
            Regex regex = new Regex(@"^Q[0-9]+$", RegexOptions.Compiled);

            // Collect worksheets that do NOT match the pattern.
            var sheetsToRemove = new System.Collections.Generic.List<Worksheet>();
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (!regex.IsMatch(sheet.Name))
                {
                    sheetsToRemove.Add(sheet);
                }
            }

            // Remove the non‑matching worksheets.
            foreach (Worksheet sheet in sheetsToRemove)
            {
                workbook.Worksheets.RemoveAt(sheet.Index);
            }

            // Save the filtered workbook to a new file.
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Filtered workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
