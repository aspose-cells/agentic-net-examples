// Title: How to read the IsAutomaticPaperSize flag for every worksheet in an Excel workbook using Aspose.Cells for .NET and store the results in a dictionary
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, iterates through all worksheets, extracts the PageSetup.IsAutomaticPaperSize property, and returns a Dictionary<string,bool> mapping worksheet names to their automatic paper size setting. | Generate a method that takes a Workbook instance and produces a Dictionary where each key is a worksheet name and each value is the IsAutomaticPaperSize boolean from that sheet's PageSetup.
// Common Searches: Aspose.Cells .NET get automatic paper size flag for each worksheet in a workbook | C# read PageSetup.IsAutomaticPaperSize for all sheets using Aspose.Cells | store worksheet name and IsAutomaticPaperSize in a dictionary Aspose.Cells | enumerate worksheets and retrieve printing settings with Aspose.Cells
// Tags: read worksheet page setup IsAutomaticPaperSize Aspose.Cells | dictionary mapping worksheet names to printing flags .NET | iterate workbook worksheets Aspose.Cells C# | extract sheet-level printing settings Excel Aspose.Cells | load Excel file and access PageSetup properties Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, loops through each worksheet, reads the PageSetup.IsAutomaticPaperSize property, stores the boolean values in a Dictionary keyed by worksheet name, and prints the results.
class Program
{
    static void Main()
    {
        // Path to the Excel file
        string filePath = "input.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(filePath);

        // Dictionary to hold worksheet name and its IsAutomaticPaperSize value
        Dictionary<string, bool> automaticPaperSizeMap = new Dictionary<string, bool>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Retrieve the IsAutomaticPaperSize property from the sheet's PageSetup
            bool isAuto = sheet.PageSetup.IsAutomaticPaperSize;

            // Store the result in the dictionary using the sheet name as the key
            automaticPaperSizeMap[sheet.Name] = isAuto;
        }

        // Example usage: print the results
        foreach (var kvp in automaticPaperSizeMap)
        {
            Console.WriteLine($"Worksheet: {kvp.Key}, IsAutomaticPaperSize: {kvp.Value}");
        }
    }
}
