// Title: Set workbook CultureInfo and refresh pivot tables to create multilingual labels with Aspose.Cells for .NET
// Common Searches: how to change workbook culture to French before refreshing pivot tables using Aspose.Cells C# | Aspose.Cells example for applying different CultureInfo to multiple Excel files and updating pivot tables | C# code to localize pivot table field captions by setting workbook Settings.CultureInfo | refresh pivot tables after changing workbook globalization settings in Aspose.Cells | generate German and French pivot table labels from an English source workbook with Aspose.Cells
// Tags: set workbook cultureinfo Aspose.Cells | pivot table data refresh Aspose.Cells | multilingual pivot labels .NET | process multiple workbooks distinct cultures | apply globalization settings Excel Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // For PivotTable class

// The sample loads an English workbook, assigns a specific CultureInfo (e.g., fr-FR or de-DE) to the workbook's Settings, iterates through each worksheet to refresh and calculate all pivot tables, and saves the result as a localized file. The ProcessWorkbook method is called separately for each target language, enabling distinct globalization settings per workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Process English source workbook for French localization
            ProcessWorkbook("Workbook_English.xlsx", "Workbook_French.xlsx", new CultureInfo("fr-FR"));
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing French workbook: {ex.Message}");
        }

        try
        {
            // Process English source workbook for German localization
            ProcessWorkbook("Workbook_English.xlsx", "Workbook_German.xlsx", new CultureInfo("de-DE"));
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing German workbook: {ex.Message}");
        }
    }

    /// <param name="inputPath">Path to the source workbook.</param>
    /// <param name="outputPath">Path where the localized workbook will be saved.</param>
    /// <param name="culture">CultureInfo to apply.</param>
    private static void ProcessWorkbook(string inputPath, string outputPath, CultureInfo culture)
    {
        if (!File.Exists(inputPath))
            throw new FileNotFoundException($"Input file not found: {inputPath}");

        // Load the workbook
        Workbook workbook = new Workbook(inputPath);

        // Apply the desired culture to the workbook settings
        workbook.Settings.CultureInfo = culture;

        // Refresh all pivot tables to reflect the new culture
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            foreach (PivotTable pivot in sheet.PivotTables)
            {
                pivot.RefreshData();
                pivot.CalculateData();
            }
        }

        // Save the localized workbook
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved: {outputPath}");
    }
}
