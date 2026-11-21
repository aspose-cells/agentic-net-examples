// Title: Apply a Brazilian Portuguese custom percentage format (comma decimal) to a cell range using Aspose.Cells for .NET
// Common Searches: how to format percentages with a comma decimal separator in Aspose.Cells C# | set pt-BR culture and apply custom number format to Excel range using Aspose.Cells | apply only number format to a cell range with StyleFlag in Aspose.Cells | Aspose.Cells custom number format for Brazilian Portuguese locale | C# load workbook and format B2:B10 as 0,00% using Aspose.Cells
// Tags: custom percentage format with comma Aspose.Cells | Brazilian Portuguese pt-BR number format Excel | apply number format to range Aspose.Cells | StyleFlag NumberFormat only Aspose.Cells | culture-specific workbook formatting Aspose.Cells | custom number format string 0,00% Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example sets the current culture to pt-BR, loads an existing workbook, creates a style with the custom number format "0,00%" (Brazilian percentage format using a comma), applies this style only to the number format of cells B2:B10 via a StyleFlag, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Apply Brazilian Portuguese culture (pt-BR)
            CultureInfo ptBR = new CultureInfo("pt-BR");
            CultureInfo.CurrentCulture = ptBR;

            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Prevent FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range that contains percentage values
            Aspose.Cells.Range percentRange = sheet.Cells.CreateRange("B2:B10");

            // Create a style with a custom number format for percentages
            Style percentStyle = workbook.CreateStyle();
            // Brazilian format uses comma as decimal separator
            percentStyle.Custom = "0,00%";

            // Apply only the number format to the range
            StyleFlag flag = new StyleFlag { NumberFormat = true };
            percentRange.ApplyStyle(percentStyle, flag);

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
