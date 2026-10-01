// Title: Programmatically hide PivotTable UI ribbons in an Excel file with Aspose.Cells for .NET before saving
// AI Prompts: Write C# code using Aspose.Cells that loads an existing .xlsx file, turns off the PivotTable UI ribbons, and saves the workbook to a new location. | Modify the provided Aspose.Cells example to include a check for the EnablePivotTableUI property and set it to false when the API supports it, while preserving error handling. | Create a robust C# routine that verifies the input workbook, suppresses PivotTable toolbars using Aspose.Cells settings, and writes the result to an output path.
// Common Searches: Aspose.Cells hide pivot table ribbon before saving workbook C# | remove PivotTable toolbar from Excel file with Aspose.Cells in C# | C# program to turn off PivotTable toolbar in generated Excel file | how to suppress PivotTable UI in Aspose.Cells when exporting | Excel pivot table UI removal using Aspose.Cells .NET
// Tags: Aspose.Cells disable pivot table UI | C# suppress Excel pivot table ribbons | Aspose.Cells UI ribbon suppression | save workbook without UI clutter Aspose | Excel UI clutter reduction .NET

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates loading an existing Excel workbook with Aspose.Cells, optionally disabling the PivotTable UI ribbons via the Settings.EnablePivotTableUI property (commented because not supported in the current version), and saving the workbook to a new file with basic file‑existence validation and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // The EnablePivotTableUI property is not available in the current Aspose.Cells version.
            // If needed, this line can be uncommented when the property becomes supported.
            // workbook.Settings.EnablePivotTableUI = false;

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
