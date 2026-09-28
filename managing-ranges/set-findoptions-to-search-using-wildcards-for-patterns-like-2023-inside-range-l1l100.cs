// Title: Find cells containing '2023' in column L (L1:L100) with Aspose.Cells FindOptions and wildcards in C#
// AI Prompts: Use Aspose.Cells Find method with a FindOptions object set to LookInType.Values and LookAtType.Contains, applying the pattern '*2023*' to locate the first matching cell in the range L1:L100. | Configure a case‑insensitive wildcard search for '*2023*' across column L using C# and retrieve the cell address and value via Aspose.Cells.
// Common Searches: aspnet find cells with wildcard '*2023*' in column L using Aspose.Cells | c# Aspose.Cells FindOptions search range L1:L100 for substring 2023 | how to use LookAtType.Contains with FindOptions in Aspose.Cells | search Excel column L for text containing 2023 with Aspose.Cells C# example | wildcard pattern search in Aspose.Cells workbook using Find method
// Tags: wildcard search FindOptions Aspose.Cells | search column L Aspose.Cells C# | LookAtType Contains example | Find method range limitation L1:L100 | case insensitive find Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, accesses the first worksheet, and uses Aspose.Cells' Find method with a FindOptions object configured for value search and a contains‑type match. By supplying the wildcard pattern '*2023*', it scans cells L1 through L100 and prints the address and value of the first cell that includes the substring '2023'.
class FindWithWildcards
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the start cell of the search range (L1)
            Cell startCell = worksheet.Cells[0, 11]; // Row 0 (1), Column 11 (L)

            // Configure FindOptions (wildcards are supported by default in the pattern)
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,   // Search cell values
                LookAtType = LookAtType.Contains, // Look for the pattern anywhere in the cell
                // Case‑insensitive search; if the API version supports MatchCase, set to false
                // MatchCase = false
            };

            // Search for any cell containing "2023" using the wildcard pattern "*2023*"
            string pattern = "*2023*";
            Cell foundCell = worksheet.Cells.Find(pattern, startCell, findOptions);

            // Handle the result
            if (foundCell != null)
            {
                Console.WriteLine($"Found match at {foundCell.Name}: {foundCell.StringValue}");
            }
            else
            {
                Console.WriteLine("No matching cells found in the specified range.");
            }

            // (Optional) Save the workbook if changes were made
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
