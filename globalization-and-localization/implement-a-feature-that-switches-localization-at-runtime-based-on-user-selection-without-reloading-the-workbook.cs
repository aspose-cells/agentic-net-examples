// Title: Change Excel workbook culture at runtime with Aspose.Cells for .NET without reloading the file
// Common Searches: Aspose.Cells C# change workbook culture after loading | Set Workbook.Settings.CultureInfo dynamically in .NET | Recalculate formulas after changing culture in Aspose.Cells | Switch Excel localization to French or German without reopening file | Runtime localization of Excel files using Aspose.Cells API | Apply different CultureInfo to the same workbook in C#
// Tags: Aspose.Cells workbook culture change | Settings.CultureInfo property usage | formula recalculation after culture update | dynamic Excel localization .NET | apply CultureInfo without reloading workbook | Aspose.Cells localization support

using Aspose.Cells;
using System;
using System.Globalization;
using System.IO;

// The WorkbookLocalizationManager loads an Excel workbook, lets you change its culture at runtime by assigning a new CultureInfo to Settings.CultureInfo, recalculates formulas to reflect culture‑specific formatting, and saves the file—all without reloading the workbook.
public class WorkbookLocalizationManager
{
    private Workbook _workbook;

    // Load the workbook using the provided load rule
    public WorkbookLocalizationManager(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"Input file not found: {filePath}");

        // {LoadWorkbook}
        _workbook = new Workbook(filePath);
    }

    // Switches the workbook's localization at runtime without reloading
    public void SwitchLocalization(string cultureName)
    {
        // Apply the selected culture to the workbook settings
        _workbook.Settings.CultureInfo = new CultureInfo(cultureName);
        // Recalculate formulas so that culture‑specific formatting takes effect
        _workbook.CalculateFormula();
    }

    // Save the workbook using the provided save rule
    public void Save(string outputPath)
    {
        // {SaveWorkbook}
        _workbook.Save(outputPath);
    }

    // Example usage
    public static void Example()
    {
        const string inputPath = "input.xlsx";

        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' does not exist.");
            return;
        }

        var manager = new WorkbookLocalizationManager(inputPath);

        try
        {
            // User selects French localization
            manager.SwitchLocalization("fr-FR");
            manager.Save("output_fr.xlsx");
            Console.WriteLine("Saved French localized workbook as output_fr.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during French localization: {ex.Message}");
        }

        try
        {
            // User selects German localization
            manager.SwitchLocalization("de-DE");
            manager.Save("output_de.xlsx");
            Console.WriteLine("Saved German localized workbook as output_de.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during German localization: {ex.Message}");
        }
    }

    // Entry point for the application
    public static void Main(string[] args)
    {
        try
        {
            Example();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
