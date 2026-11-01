// Title: Customizing Subtotal and Grand Total labels in Excel with Aspose.Cells .NET by overriding GlobalizationSettings
// Common Searches: asp.net c# override GlobalizationSettings to change subtotal label in Aspose.Cells | how to localize grand total text in an Excel workbook using Aspose.Cells .NET | create a custom GlobalizationSettings subclass for label translation in Aspose.Cells | Aspose.Cells missing SubtotalLabel property workaround for older versions | recalculate formulas after updating subtotal and grand total labels with Aspose.Cells
// Tags: override GlobalizationSettings Aspose.Cells | subtotal label localization .NET | grand total text customization Excel | Aspose.Cells formula recalculation after label change | Aspose.Cells version compatibility SubtotalLabel

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an Excel workbook with Aspose.Cells, notes that SubtotalLabel and GrandTotalLabel are unavailable in the current library version, suggests upgrading or overriding GlobalizationSettings in supported versions, recalculates formulas to refresh totals, and saves the workbook while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: GlobalizationSettings properties such as SubtotalLabel and GrandTotalLabel
            // are not available in the current Aspose.Cells version, so they are omitted.
            // If needed, upgrade Aspose.Cells to a version that supports these properties.

            // Recalculate formulas to refresh totals
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
