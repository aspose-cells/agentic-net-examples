// Title: Filter worksheets by regex when loading an Excel workbook with Aspose.Cells for .NET (retain only Q‑number sheets)
// Common Searches: asp.net filter Excel worksheets by name pattern using Aspose.Cells | c# load workbook and keep only sheets that match ^Q\d+$ with Aspose.Cells | how to use LoadOptions.LoadFilter to include only worksheets named Q1 Q2 in Aspose.Cells | remove Excel sheets that do not follow naming convention Q<number> using Aspose.Cells .NET | regex based worksheet selection during workbook load Aspose.Cells C#
// Tags: regex worksheet filtering Aspose.Cells .NET | load workbook with sheet name pattern Aspose.Cells | retain specific worksheets Excel C# | exclude nonmatching sheets Aspose.Cells | worksheet name pattern Q\d+ filter

using Aspose.Cells;
using System;
using System.IO;
using System.Text.RegularExpressions;

// The example loads an Excel file with Aspose.Cells, defines a regular expression ^Q\d+$ to identify worksheets named like Q1, Q23, etc., removes all sheets that do not match this pattern, and saves the resulting workbook. It demonstrates how to filter worksheets by name using regex in C#, handling missing files and runtime errors.
class Program
{
    static void Main()
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
            // Load the workbook without any filter
            Workbook workbook = new Workbook(inputPath);

            // Define the regex pattern for worksheet names to keep (e.g., Q1, Q23, etc.)
            string pattern = @"^Q\d+$";
            Regex regex = new Regex(pattern, RegexOptions.Compiled);

            // Collect indexes of worksheets that do NOT match the pattern
            var indexesToRemove = new System.Collections.Generic.List<int>();
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                if (!regex.IsMatch(sheet.Name))
                {
                    indexesToRemove.Add(i);
                }
            }

            // Remove worksheets in descending order to preserve indexes
            for (int i = indexesToRemove.Count - 1; i >= 0; i--)
            {
                workbook.Worksheets.RemoveAt(indexesToRemove[i]);
            }

            // Save the filtered workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Filtered workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
