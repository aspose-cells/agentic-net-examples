// Title: C# example: Load an Excel workbook and keep only worksheets whose names match the pattern Q[0-9]+ using a regular expression filter
// Common Searches: how to filter worksheets by name pattern when opening an Excel file with Aspose.Cells in C# | C# Aspose.Cells load workbook and retain only sheets starting with Q followed by numbers | using regex to select specific worksheets during workbook processing with Aspose.Cells | remove non‑matching worksheets from an Excel workbook programmatically in C# | Aspose.Cells example for keeping only Q1, Q2, etc. sheets after loading
// Tags: Aspose.Cells worksheet name regex filter C# | load workbook selective sheets Aspose.Cells | regex based worksheet selection Excel C# | filter Excel worksheets by pattern Aspose.Cells | remove nonmatching sheets Aspose.Cells C#

using Aspose.Cells;
using System;
using System.IO;
using System.Text.RegularExpressions;

// The sample checks that the source Excel file exists, opens it with Aspose.Cells, iterates through the worksheets in reverse order, removes any sheet whose name does not satisfy the regular expression ^Q\d+$, prints the names of the retained sheets, and saves the resulting workbook to a new file while handling exceptions.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                return;
            }

            // Load the workbook without a filter (filtering will be done manually)
            Workbook workbook = new Workbook(inputFile);

            // Remove worksheets whose names do not match the pattern Q[0-9]+
            for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
            {
                Worksheet sheet = workbook.Worksheets[i];
                if (!Regex.IsMatch(sheet.Name, @"^Q\d+$"))
                {
                    workbook.Worksheets.RemoveAt(i);
                }
            }

            // Output the names of the remaining worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine(sheet.Name);
            }

            // Save the processed workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
