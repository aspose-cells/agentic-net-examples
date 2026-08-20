// Title: Batch replace domain URLs in a named range using Aspose.Cells for .NET (C#)
// Description: Load an Excel workbook, locate the named range "Links", iterate through its string cells, replace every occurrence of an old domain with a new one, and save the updated file—all with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# replace text | named range URL update | batch replace domain Excel | C# Excel string replace Aspose | update hyperlinks programmatically | .NET Excel bulk edit | Aspose.Cells named range manipulation | Excel URL migration C#
// Common Searches: How to change a domain in all cells of a named range with Aspose.Cells | C# code to batch replace URLs in Excel named range | Aspose.Cells get and edit named range values | Replace old website links with new ones in an Excel file using .NET | Bulk update hyperlinks in Excel via Aspose.Cells
// Developer Intent: Replace every occurrence of a specified old domain with a new domain inside the string cells of the named range "Links".
// Use Cases: Migrate marketing URLs after a domain rebrand across a template workbook. | Swap placeholder test links with production URLs in automated report generation. | Correct outdated hyperlinks in financial models without manual editing.
// AI Prompts: Write C# code that uses Aspose.Cells to find the named range "Links" in an Excel file and replace all instances of "old.example.com" with "new.example.com". | Explain step‑by‑step how to retrieve a named range reference, parse its sheet and address, create a Range object, and perform string replacement on its cells with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace ReplaceUrlsInNamedRange
{
    // Load an Excel workbook, locate the named range "Links", iterate through its string cells, replace every occurrence of an old domain with a new one, and save the updated file—all with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "Input.xlsx";
                const string outputPath = "Output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Named range and domain settings
                const string namedRangeName = "Links";
                const string oldDomain = "old.example.com";
                const string newDomain = "new.example.com";

                // Retrieve the named range
                Name namedRange = workbook.Worksheets.Names[namedRangeName];
                if (namedRange == null)
                {
                    Console.WriteLine($"Named range \"{namedRangeName}\" not found.");
                    return;
                }

                // Get the reference string without leading '='
                string refersTo = namedRange.GetRefersTo(false, false);
                if (refersTo.StartsWith("="))
                    refersTo = refersTo.Substring(1);

                // Split into sheet name and range address
                string[] parts = refersTo.Split('!');
                if (parts.Length != 2)
                {
                    Console.WriteLine("Unable to parse the range reference.");
                    return;
                }

                string sheetName = parts[0];
                string rangeAddress = parts[1];

                // Get the worksheet
                Worksheet sheet = workbook.Worksheets[sheetName];
                if (sheet == null)
                {
                    Console.WriteLine($"Worksheet \"{sheetName}\" not found.");
                    return;
                }

                // Create the Aspose.Cells.Range explicitly to avoid ambiguity
                Aspose.Cells.Range range = sheet.Cells.CreateRange(rangeAddress);

                // Replace old domain with new domain in string cells
                foreach (Cell cell in range)
                {
                    if (cell.Type == CellValueType.IsString)
                    {
                        string original = cell.StringValue;
                        if (!string.IsNullOrEmpty(original) && original.Contains(oldDomain))
                        {
                            string updated = original.Replace(oldDomain, newDomain);
                            cell.PutValue(updated);
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
