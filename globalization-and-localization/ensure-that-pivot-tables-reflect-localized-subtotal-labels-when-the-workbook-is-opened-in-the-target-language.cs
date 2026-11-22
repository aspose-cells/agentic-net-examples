// Title: Localize pivot table subtotal captions in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set pivot table subtotal caption for German locale | how to change SubtotalCaption of all pivot tables using Aspose.Cells .NET | localize pivot table labels in Excel workbook with CultureInfo in Aspose.Cells | apply language‑specific subtotal text to Excel pivot tables programmatically | Aspose.Cells reflection example to update PivotTable properties
// Tags: set SubtotalCaption property Aspose.Cells | configure workbook CultureInfo Aspose.Cells | reflection based PivotTable property update C# | Excel pivot tables language localization | Aspose.Cells .NET workbook internationalization

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable class

// The example loads an existing Excel file, assigns a target CultureInfo (e.g., German) to the workbook, iterates through each worksheet and its pivot tables, and—when the SubtotalCaption property is available—sets it to a culture‑specific label via reflection. The modified workbook is then saved to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (rule: load-workbook)
            Workbook workbook = new Workbook(inputPath);

            // Set the target culture for localization (e.g., German)
            workbook.Settings.CultureInfo = new CultureInfo("de-DE");

            // Iterate through all worksheets and their pivot tables
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    try
                    {
                        // Aspose.Cells versions prior to 23.x do not expose SubtotalCaption.
                        // If the property exists, set a localized caption; otherwise, skip.
                        var prop = typeof(PivotTable).GetProperty("SubtotalCaption");
                        if (prop != null && prop.CanWrite)
                        {
                            prop.SetValue(pivot, GetLocalizedSubtotal(workbook.Settings.CultureInfo));
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log but continue processing other pivot tables
                        Console.WriteLine($"Warning: Unable to set SubtotalCaption for pivot '{pivot.Name}'. {ex.Message}");
                    }
                }
            }

            // Save the workbook (rule: save-workbook)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to return the appropriate subtotal label for a given culture
    static string GetLocalizedSubtotal(CultureInfo culture)
    {
        return culture.Name switch
        {
            "de-DE" => "Zwischensumme",
            "fr-FR" => "Sous‑total",
            "es-ES" => "Subtotal",
            "ja-JP" => "小計",
            _ => "Subtotal",
        };
    }
}
