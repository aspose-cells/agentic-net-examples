// Title: Perform a case‑insensitive find‑and‑replace of the phrase “total revenue” with “Revenue Total” in an Excel workbook using Aspose.Cells for .NET
// Common Searches: asp.net c# find and replace text in Excel workbook case insensitive Aspose.Cells | replace all occurrences of 'total revenue' in .xlsx with Aspose.Cells C# | how to search cell values containing a phrase and update them using Aspose.Cells | using FindOptions to perform contains search in Excel sheets with Aspose.Cells | regex based case‑insensitive text replacement in Excel cells C# Aspose
// Tags: case-insensitive find replace Aspose.Cells | Excel workbook text replacement C# | FindOptions contains search Aspose.Cells | regex replace cell values Aspose.Cells | iterate worksheets update text Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an existing .xlsx file with Aspose.Cells, configures FindOptions for a case‑insensitive, contains search on cell values, iterates every worksheet, finds each cell containing the phrase “total revenue”, replaces it with “Revenue Total” using a Regex replace that ignores case, writes the updated string back to the cell, and saves the modified workbook as a new file while handling missing input and runtime errors.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Text to find (case‑insensitive) and its replacement
            string searchText = "total revenue";
            string replaceText = "Revenue Total";

            // Configure find options for a case‑insensitive, contains search in cell values
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,
                LookAtType = LookAtType.Contains,
                CaseSensitive = false
            };

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Find the first matching cell
                Cell cell = sheet.Cells.Find(searchText, null, findOptions);

                // Loop through all matches
                while (cell != null)
                {
                    // Perform a case‑insensitive replacement within the cell's string value
                    string originalValue = cell.StringValue;
                    string newValue = Regex.Replace(
                        originalValue,
                        Regex.Escape(searchText),
                        replaceText,
                        RegexOptions.IgnoreCase);

                    // Write the updated value back to the cell
                    cell.PutValue(newValue);

                    // Find the next matching cell starting after the current one
                    cell = sheet.Cells.Find(searchText, cell, findOptions);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
