// Title: Programmatically hide PivotTable UI ribbons in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells hide pivot table ribbon before saving workbook C# | disable pivot table UI in Excel file with Aspose.Cells .NET | remove pivot table toolbar programmatically using Aspose.Cells | how to suppress pivot table ribbon in generated Excel via Aspose.Cells | Aspose.Cells workbook settings to turn off pivot table UI
// Tags: pivot table UI suppression Aspose.Cells | disable pivot table ribbon .NET | hide Excel pivot table UI programmatically | Aspose.Cells workbook settings for UI | reduce Excel UI clutter Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook with Aspose.Cells, notes that the current library version does not provide a direct property to hide the PivotTable UI ribbon, includes a commented placeholder for a future EnablePivotTableUI setting, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // The property to hide PivotTable UI is not available in the current Aspose.Cells version.
            // If a future version provides it, uncomment the line below:
            // workbook.Settings.EnablePivotTableUI = false;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
