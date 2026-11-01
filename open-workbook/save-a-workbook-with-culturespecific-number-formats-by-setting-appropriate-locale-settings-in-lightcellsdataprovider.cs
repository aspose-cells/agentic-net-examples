// Title: How to save an Excel workbook with French‑style number formatting using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set workbook culture to French for comma decimal separator | Save Excel file with locale‑specific number format using Aspose.Cells .NET | Apply French number format to cells before saving workbook in Aspose.Cells | Configure CultureInfo in Aspose.Cells to generate French formatted numbers
// Tags: Aspose.Cells set workbook CultureInfo | French locale number format Excel C# | save workbook with locale specific formatting Aspose.Cells | apply built‑in number style Aspose.Cells | SaveFormat.Xlsx culture aware Aspose.Cells

using System;
using System.Data;
using System.Globalization;
using Aspose.Cells;

// The example creates a new Workbook, assigns Settings.CultureInfo = new CultureInfo("fr-FR") so numbers use a comma as the decimal separator, writes three numeric values, applies a built‑in style with two decimal places, and saves the file as CultureSpecificNumbers.xlsx, producing French‑style number formatting in the resulting Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set workbook culture to French (comma as decimal separator)
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Fill some numeric data
            sheet.Cells["A1"].PutValue(1234.56);
            sheet.Cells["A2"].PutValue(7890.12);
            sheet.Cells["A3"].PutValue(3456.78);

            // Apply a number format (2 decimal places)
            Style style = workbook.CreateStyle();
            style.Number = 2; // Built‑in format with 2 decimal places
            sheet.Cells["A1"].SetStyle(style);
            sheet.Cells["A2"].SetStyle(style);
            sheet.Cells["A3"].SetStyle(style);

            // Save the workbook; number formats will reflect the French culture
            string outputPath = "CultureSpecificNumbers.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
