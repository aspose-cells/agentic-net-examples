// Title: Perform a case‑insensitive search for the word “invoice” in all workbook‑scoped named ranges using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells find cells containing 'invoice' in workbook scoped named ranges | search case-insensitive text in Excel named ranges using Aspose.Cells .NET | enumerate all workbook‑level named ranges and locate a specific word with Aspose.Cells | how to retrieve cell addresses of a keyword inside named ranges in a .xlsx file using C# | Aspose.Cells iterate named ranges and check for text match ignoring case
// Tags: case-insensitive text search in Aspose.Cells named ranges | enumerate workbook scoped names Aspose.Cells .NET | retrieve cell addresses from named ranges C# | search keyword in Excel named ranges using Aspose.Cells | Aspose.Cells iterate over named ranges for text match

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel workbook, iterates through every workbook‑scoped named range, examines each cell within those ranges, and records the addresses of cells whose text contains the word “invoice” using a case‑insensitive comparison, then outputs the named range name together with each matching cell address.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook;

        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // List to hold the addresses where the word "invoice" is found
        List<string> foundAddresses = new List<string>();

        try
        {
            // Iterate through all workbook‑scoped named ranges
            foreach (Name namedRange in workbook.Worksheets.Names)
            {
                // Get the range that the name refers to
                AsposeRange range = namedRange.GetRange();

                // Loop through each cell in the range
                for (int row = range.FirstRow; row <= range.FirstRow + range.RowCount - 1; row++)
                {
                    for (int col = range.FirstColumn; col <= range.FirstColumn + range.ColumnCount - 1; col++)
                    {
                        Cell cell = range.Worksheet.Cells[row, col];
                        string cellText = cell.StringValue;

                        // Perform a case‑insensitive search for the word "invoice"
                        if (!string.IsNullOrEmpty(cellText) &&
                            cellText.IndexOf("invoice", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Record the address (e.g., Sheet1!A1) along with the named range name
                            string address = $"{range.Worksheet.Name}!{cell.Name}";
                            foundAddresses.Add($"{namedRange.Text} -> {address}");
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while processing named ranges: {ex.Message}");
            return;
        }

        // Output the results
        Console.WriteLine("Occurrences of the word \"invoice\" in workbook‑scoped named ranges:");
        foreach (string entry in foundAddresses)
        {
            Console.WriteLine(entry);
        }
    }
}
