// Title: Apply a custom currency number format to all cells in the "CurrencyValues" named range with Aspose.Cells for .NET
// Common Searches: how to set a custom currency number format for a named range using Aspose.Cells in C# | Aspose.Cells C# apply number format to every cell in a specific named range | using StyleFlag to change only the number format of a range in Aspose.Cells | format cells in the "CurrencyValues" range as currency with Aspose.Cells .NET | apply a custom Excel currency format to a named range programmatically
// Tags: custom currency format Aspose.Cells | named range number format C# | StyleFlag NumberFormat Aspose.Cells | range.ApplyStyle currency style | Excel named range formatting .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, retrieves the named range "CurrencyValues", creates a style with the custom currency format "$#,##0.00", applies only the number format to the entire range using a StyleFlag, and saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
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

        // Retrieve the named range "CurrencyValues"
        Aspose.Cells.Range range = workbook.Worksheets.GetRangeByName("CurrencyValues");
        if (range == null)
        {
            Console.WriteLine("Named range 'CurrencyValues' not found.");
            return;
        }

        // Create a style with custom currency number format
        Style currencyStyle = workbook.CreateStyle();
        currencyStyle.Custom = "$#,##0.00";

        // Apply only the number format to the entire range
        StyleFlag flag = new StyleFlag { NumberFormat = true };
        range.ApplyStyle(currencyStyle, flag);

        try
        {
            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
