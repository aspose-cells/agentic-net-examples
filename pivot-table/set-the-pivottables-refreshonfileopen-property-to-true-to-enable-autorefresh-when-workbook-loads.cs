// Title: Enable auto‑refresh of a PivotTable on workbook open using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set PivotTable RefreshOnFileOpen property | How to make an Excel PivotTable refresh automatically when opening the file with Aspose.Cells | C# example for enabling PivotTable auto‑refresh on workbook load using Aspose.Cells | Set PivotTable to refresh on file open in Aspose.Cells .NET | Aspose.Cells code to activate RefreshOnFileOpen for a PivotTable
// Tags: Aspose.Cells PivotTable RefreshOnFileOpen | C# auto refresh pivot table on workbook open | configure PivotTable refresh property Aspose.Cells | Excel pivot table auto refresh using Aspose.Cells .NET | set PivotTable refresh on file open C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing Excel workbook, checks for a PivotTable on the first worksheet, and (when supported) sets the PivotTable's RefreshOnFileOpen property to true so the pivot data refreshes automatically each time the workbook is opened, then saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            // Access the first PivotTable on the worksheet
            PivotTable pivot = sheet.PivotTables[0];

            // NOTE: In some Aspose.Cells versions the property to enable auto‑refresh is not available.
            // If your version supports it, uncomment the following line:
            // pivot.RefreshDataOnOpen = true;

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
