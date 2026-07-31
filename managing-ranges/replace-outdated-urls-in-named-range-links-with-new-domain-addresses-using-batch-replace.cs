// Title: Replace old domain URLs in the named range "Links" using Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, accesses the named range "Links", scans each string cell, swaps the legacy domain (http://old.example.com) with the new domain (https://new.example.com), and saves the updated file.
// Keywords: Aspose.Cells | C# | replace text in Excel | named range | URL batch update | hyperlink migration | workbook cell iteration | domain change
// Common Searches: Aspose.Cells replace URLs in a named range | C# batch update Excel hyperlinks | How to change domain in Excel cells using Aspose | Replace old website links in Excel with new domain
// Developer Intent: Update every occurrence of an outdated domain within the "Links" named range of an Excel workbook using Aspose.Cells.
// Use Cases: Migrate all hyperlinks after a website rebranding by updating the domain in a predefined named range. | Prepare a report for external distribution by converting internal URLs to public‑facing ones in bulk. | Fix broken links in a financial model where every source URL shares the same obsolete domain.
// AI Prompts: Generate C# code with Aspose.Cells that replaces a substring in all string cells of a specified named range, handling missing ranges gracefully. | Provide an Aspose.Cells snippet that validates the existence of a named range before performing a batch URL replacement and then saves the workbook. | Write a reusable C# method that accepts old and new domain strings and updates every matching URL within a given named range using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsReplaceInNamedRange
{
    // Loads an Excel workbook, accesses the named range "Links", scans each string cell, swaps the legacy domain (http://old.example.com) with the new domain (https://new.example.com), and saves the updated file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "InputWorkbook.xlsx";
                const string outputPath = "OutputWorkbook.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Define the old and new domain strings to be replaced
                string oldDomain = "http://old.example.com";
                string newDomain = "https://new.example.com";

                // Retrieve the named range "Links"
                AsposeRange linksRange = workbook.Worksheets.GetRangeByName("Links");

                // If the named range does not exist, exit gracefully
                if (linksRange == null)
                {
                    Console.WriteLine("Named range 'Links' not found.");
                    return;
                }

                // Iterate through each cell in the range and replace occurrences of the old domain
                foreach (Cell cell in linksRange)
                {
                    // Process only cells that contain string values
                    if (cell.Type == CellValueType.IsString)
                    {
                        string originalText = cell.StringValue;
                        if (!string.IsNullOrEmpty(originalText) && originalText.Contains(oldDomain))
                        {
                            string updatedText = originalText.Replace(oldDomain, newDomain);
                            cell.PutValue(updatedText);
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
