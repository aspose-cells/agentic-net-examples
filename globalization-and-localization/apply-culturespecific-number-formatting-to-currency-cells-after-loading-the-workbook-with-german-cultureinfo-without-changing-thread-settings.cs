// Title: Apply German (de-DE) currency number format to a cell range in an Excel workbook using Aspose.Cells for .NET without changing thread culture
// Common Searches: how to format cells as Euro currency with German locale in Aspose.Cells C# | set workbook culture to de-DE for number formatting in Aspose.Cells without modifying thread culture | apply custom number format '#,##0.00 €' to a range of cells using Aspose.Cells .NET | Aspose.Cells C# apply German currency format to B2:B10 range | use Workbook.Settings.CultureInfo to display Excel numbers in German currency format
// Tags: Workbook.Settings.CultureInfo de-DE | custom number format Euro Aspose.Cells | StyleFlag NumberFormat range B2:B10 | culture-specific currency formatting Excel .NET | German locale Excel workbook Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file (or creates a new workbook), assigns the German culture (de-DE) to the workbook via Workbook.Settings.CultureInfo, defines a custom number format "#,##0.00 €", and applies this format only to the number format of cells B2:B10 using a StyleFlag. The workbook is then saved, preserving the German currency display without altering the application's thread culture.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook or create a new one if the file is missing
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Apply German culture to the workbook
            workbook.Settings.CultureInfo = new CultureInfo("de-DE");

            // Create a style with German currency format
            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Custom = "#,##0.00 €";

            // Define the range B2:B10 on the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Aspose.Cells.Range currencyRange = sheet.Cells.CreateRange("B2:B10");

            // Apply the currency style (only number format flag)
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;
            currencyRange.ApplyStyle(currencyStyle, flag);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
