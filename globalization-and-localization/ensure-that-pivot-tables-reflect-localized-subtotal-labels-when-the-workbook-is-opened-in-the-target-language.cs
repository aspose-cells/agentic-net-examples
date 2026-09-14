// Title: Set workbook CultureInfo to localize pivot table subtotal labels in Excel using Aspose.Cells for .NET
// AI Prompts: Load an .xlsx file with Aspose.Cells, assign a specific CultureInfo (e.g., fr-FR) to Workbook.Settings.CultureInfo, and save the file so pivot tables display subtotal captions in the target language. | Iterate through all worksheets in a workbook, ensure each pivot table inherits the workbook’s CultureInfo, and verify that the localized labels appear when the file is opened.
// Common Searches: how to change pivot table language to French with Aspose.Cells C# | Aspose.Cells set workbook culture for localized subtotal captions | C# programmatically set Excel file language for pivot tables | localize Excel pivot table labels using CultureInfo in .NET | Aspose.Cells Workbook.Settings.CultureInfo effect on pivot table captions
// Tags: Aspose.Cells workbook language configuration | pivot table subtotal caption localization | C# assign Excel workbook CultureInfo | Excel file language programmatically .NET | Aspose.Cells pivot table language setting

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Provides PivotTable type

// The example loads an existing Excel workbook, sets its Settings.CultureInfo to a target locale (e.g., fr-FR), and saves the file. Aspose.Cells automatically applies the culture to pivot tables, causing subtotal labels to appear in the localized language when the workbook is opened.
class PivotTableLocalization
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Set the workbook culture to the target language (e.g., French)
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Iterate through all worksheets and their pivot tables
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    // Note: Aspose.Cells does not expose a SubtotalCaption property.
                    // Localization of pivot table captions is handled by the workbook's CultureInfo.
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
