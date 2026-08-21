// Title: Aspose.Cells for .NET – Replace "N/A" Formula Results with Blank Cells in a Named Range
// Description: Loads an Excel workbook, forces formula calculation, locates a named range (e.g., MyRange), iterates through its cells and clears any that display the text "N/A", then saves the modified file. Ideal for cleaning up reports or preparing data for downstream processing.
// Keywords: Aspose.Cells | C# | replace N/A | blank cells | named range | formula result | clear N/A values | Excel automation | Workbook.CalculateFormula | range iteration
// Common Searches: Aspose.Cells replace N/A with empty cell in named range | C# clear N/A values from Excel named range using Aspose | How to remove N/A text returned by formulas in Aspose.Cells
// Developer Intent: Identify cells that return the string "N/A" within a specific named range and convert them to empty cells.
// Use Cases: Sanitize a financial report by removing placeholder N/A values before publishing. | Prepare data for import into another system that cannot handle the N/A string. | Automate workbook cleanup in a scheduled task that processes multiple files.
// AI Prompts: Generate C# code with Aspose.Cells that finds a named range and replaces any cell showing "N/A" with an empty string, handling missing files and missing ranges. | Create a reusable method that accepts a workbook path and a range name, clears N/A results, and returns the updated workbook.

using Aspose.Cells;
using System;
using System.IO;

// Loads an Excel workbook, forces formula calculation, locates a named range (e.g., MyRange), iterates through its cells and clears any that display the text "N/A", then saves the modified file. Ideal for cleaning up reports or preparing data for downstream processing.
class ReplaceNAWithBlank
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string rangeName = "MyRange";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Ensure formulas are calculated so we can read their results
            workbook.CalculateFormula();

            // Retrieve the named range; if it does not exist, exit after saving (or create a new workbook)
            Name namedRange = workbook.Worksheets.Names[rangeName];
            if (namedRange == null)
            {
                Console.WriteLine($"Named range \"{rangeName}\" does not exist.");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
                return;
            }

            // Get the range reference in A1 style (e.g., Sheet1!$A$1:$B$10)
            string refersTo = namedRange.GetRefersTo(false, false);
            if (refersTo.StartsWith("="))
                refersTo = refersTo.Substring(1); // Remove leading '='

            // Separate sheet name and address
            Worksheet sheet;
            string address;
            int exclPos = refersTo.IndexOf('!');
            if (exclPos >= 0)
            {
                string sheetName = refersTo.Substring(0, exclPos);
                sheet = workbook.Worksheets[sheetName];
                address = refersTo.Substring(exclPos + 1);
            }
            else
            {
                sheet = workbook.Worksheets[0];
                address = refersTo;
            }

            // Create a Range object for the address
            Aspose.Cells.Range range = sheet.Cells.CreateRange(address);

            // Iterate through each cell in the range and replace "N/A" with blank
            foreach (Cell cell in range)
            {
                if (cell.StringValue == "N/A")
                {
                    cell.PutValue(string.Empty); // Clears the formula/value
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
