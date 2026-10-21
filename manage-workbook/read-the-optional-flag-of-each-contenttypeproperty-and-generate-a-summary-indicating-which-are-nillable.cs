// Title: Create a C# program that lists each worksheet’s visibility (visible or hidden) in an Excel workbook using Aspose.Cells
// Common Searches: how to determine if a worksheet is hidden using Aspose.Cells in C# | C# Aspose.Cells iterate through worksheets and get visibility status | list all sheet names with visibility flag Aspose.Cells .NET | generate console report of worksheet visibility Aspose.Cells | check worksheet IsVisible property with Aspose.Cells C# example
// Tags: Aspose.Cells read worksheet visibility | C# enumerate Excel worksheets Aspose.Cells | Aspose.Cells worksheet IsVisible property | generate worksheet visibility report .NET | list Excel sheet names with visibility Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example verifies that input.xlsx exists, loads it into an Aspose.Cells Workbook, iterates over every worksheet, uses the IsVisible property to decide if each sheet is visible or hidden, appends a line with the sheet name and its visibility to a StringBuilder, and prints the compiled summary to the console.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook = null;
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

        // Prepare a StringBuilder to collect the summary
        StringBuilder summary = new StringBuilder();

        // Iterate over each worksheet in the workbook (valid Aspose.Cells API)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Example property: visibility (visible, hidden, very hidden)
            string visibility = sheet.IsVisible ? "visible" : "hidden";
            summary.AppendLine($"Worksheet \"{sheet.Name}\" is {visibility}.");
        }

        // Output the generated summary
        Console.WriteLine(summary.ToString());
    }
}
