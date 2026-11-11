// Title: Clear all page fields from a PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to remove page filters from a PivotTable | example code for PivotTable.PageFields.Clear in Aspose.Cells | clear pivot page fields programmatically with Aspose.Cells .NET | remove all page fields from an Excel PivotTable using Aspose.Cells C#
// Tags: Aspose.Cells clear pivot page fields C# | PivotTable PageFields.Clear method Aspose | remove pivot page filters Aspose.Cells | modify pivot layout programmatically .NET | Excel workbook save after pivot changes Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Loads an existing workbook, accesses the first PivotTable on the first worksheet, clears all its page fields with the PageFields.Clear method, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook that contains the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found on the first worksheet.");
                return;
            }

            // Access the first PivotTable on the worksheet
            PivotTable pivotTable = sheet.PivotTables[0];

            // Clear all page fields from the PivotTable
            pivotTable.PageFields.Clear();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
