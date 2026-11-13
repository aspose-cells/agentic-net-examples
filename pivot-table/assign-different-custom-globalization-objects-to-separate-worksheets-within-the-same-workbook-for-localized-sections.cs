// Title: How to apply French and German number formats to separate worksheets in an Aspose.Cells workbook using C#
// Common Searches: C# Aspose.Cells set French number format on one worksheet and German on another | apply different culture-specific number styles per sheet in Aspose.Cells .NET | save workbook with localized number formatting for each worksheet using Aspose.Cells | Aspose.Cells custom number format per worksheet example C# | create Excel file with French and German numeric formats using Aspose.Cells
// Tags: culture-specific number format worksheet Aspose.Cells | French number style Excel Aspose.Cells C# | German numeric format worksheet Aspose.Cells | per-sheet custom number format Aspose.Cells | Aspose.Cells workbook localization C#

using Aspose.Cells;
using System;
using System.Globalization;
using System.IO;

// The sample creates a workbook, formats cell A1 on the first sheet with a French number pattern "#,##0.00" and on a second sheet with a German pattern "#.##0,00", writes the same value to both cells, and saves the file as LocalizedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // First worksheet: French localization
            Worksheet frenchSheet = workbook.Worksheets[0];
            frenchSheet.Name = "FrenchSection";

            // Apply French number format to cell A1
            Style frenchStyle = workbook.CreateStyle();
            frenchStyle.Custom = "#,##0.00"; // French uses comma as decimal separator
            frenchStyle.Number = 2;          // General number format
            frenchSheet.Cells["A1"].SetStyle(frenchStyle);
            frenchSheet.Cells["A1"].PutValue(12345.67);

            // Second worksheet: German localization
            Worksheet germanSheet = workbook.Worksheets.Add("GermanSection");

            // Apply German number format to cell A1
            Style germanStyle = workbook.CreateStyle();
            germanStyle.Custom = "#.##0,00"; // German uses dot for thousands, comma for decimal
            germanStyle.Number = 2;
            germanSheet.Cells["A1"].SetStyle(germanStyle);
            germanSheet.Cells["A1"].PutValue(12345.67);

            // Determine output path and ensure directory exists
            string outputPath = "LocalizedWorkbook.xlsx";
            string fullPath = Path.GetFullPath(outputPath);
            string outputDir = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
