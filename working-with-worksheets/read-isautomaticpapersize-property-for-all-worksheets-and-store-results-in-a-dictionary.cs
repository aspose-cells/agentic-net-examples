// Title: Collect IsAutomaticPaperSize setting for every worksheet into a C# dictionary using Aspose.Cells
// Common Searches: how to get worksheet page setup IsAutomaticPaperSize value with Aspose.Cells in C# | read automatic paper size flag for all sheets in an Excel file using Aspose.Cells | store worksheet names and their IsAutomaticPaperSize property in a dictionary in .NET | Aspose.Cells iterate worksheets and retrieve page setup automatic paper size | C# example to extract IsAutomaticPaperSize from each worksheet
// Tags: Aspose.Cells worksheet page setup property extraction | read IsAutomaticPaperSize flag C# | iterate workbook worksheets Aspose.Cells | dictionary mapping worksheet name to bool | automatic paper size detection Excel Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;

// The sample loads an Excel workbook with Aspose.Cells, creates a Dictionary<string, bool>, loops through each Worksheet, reads the PageSetup.IsAutomaticPaperSize property, stores the boolean value keyed by the worksheet name, and then prints the collected name‑value pairs.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Dictionary to store worksheet name and its IsAutomaticPaperSize value
        Dictionary<string, bool> isAutomaticPaperSizeMap = new Dictionary<string, bool>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Read the IsAutomaticPaperSize property from the worksheet's PageSetup
            bool isAuto = sheet.PageSetup.IsAutomaticPaperSize;

            // Store the result in the dictionary using the worksheet name as the key
            isAutomaticPaperSizeMap[sheet.Name] = isAuto;
        }

        // Example: display the collected values
        foreach (var kvp in isAutomaticPaperSizeMap)
        {
            Console.WriteLine($"Worksheet: {kvp.Key}, IsAutomaticPaperSize: {kvp.Value}");
        }
    }
}
