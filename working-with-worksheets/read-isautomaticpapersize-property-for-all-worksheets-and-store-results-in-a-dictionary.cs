// Title: Read the IsAutomaticPaperSize setting for every worksheet and store it in a C# dictionary using Aspose.Cells
// Common Searches: C# Aspose.Cells read IsAutomaticPaperSize for each worksheet in a workbook | store worksheet page‑setup automatic paper size flag in a dictionary with Aspose.Cells .NET | iterate through all worksheets and get IsAutomaticPaperSize property using Aspose.Cells | retrieve automatic paper size setting for multiple sheets in an Excel file with Aspose.Cells | dictionary of worksheet names and IsAutomaticPaperSize values in C#
// Tags: Aspose.Cells worksheet page setup IsAutomaticPaperSize | C# dictionary mapping worksheet names to page setup flags | iterate workbook worksheets Aspose.Cells .NET | read automatic paper size setting Excel Aspose.Cells | page setup property extraction Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, loops through each worksheet, reads the IsAutomaticPaperSize flag from the worksheet's PageSetup, stores the boolean result in a Dictionary keyed by worksheet name, and then outputs the worksheet names with their corresponding flag values.
class Program
{
    static void Main()
    {
        // Path to the Excel file
        string filePath = "input.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(filePath);

        // Dictionary to store worksheet name and its IsAutomaticPaperSize value
        Dictionary<string, bool> automaticPaperSizeMap = new Dictionary<string, bool>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Read the IsAutomaticPaperSize property from the worksheet's PageSetup
            bool isAutomatic = sheet.PageSetup.IsAutomaticPaperSize;

            // Store the result in the dictionary using the worksheet name as the key
            automaticPaperSizeMap[sheet.Name] = isAutomatic;
        }

        // Example usage: print the results
        foreach (var kvp in automaticPaperSizeMap)
        {
            Console.WriteLine($"Worksheet: {kvp.Key}, IsAutomaticPaperSize: {kvp.Value}");
        }
    }
}
