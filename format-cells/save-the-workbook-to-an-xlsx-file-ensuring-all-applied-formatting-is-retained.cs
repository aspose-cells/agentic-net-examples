// Title: Save an Aspose.Cells workbook to XLSX with all cell formatting retained using C#
// AI Prompts: Write C# code that loads or creates a workbook, applies a red bold font with yellow background to the range A1:B2, and saves it as an XLSX file while preserving the applied styles. | Show how to use Aspose.Cells SaveFormat.Xlsx to export a workbook and keep cell shading and font attributes intact.
// Common Searches: Aspose.Cells C# save workbook to xlsx without losing cell styles | preserve cell background color and font when exporting Excel with Aspose.Cells | how to keep formatting of range A1:B2 after saving workbook using Aspose.Cells in C#
// Tags: Aspose.Cells workbook.Save preserving styles | SaveFormat.Xlsx retain cell formatting | apply style to range A1:B2 Aspose.Cells | C# export Excel with font and background preservation | Aspose.Cells keep cell shading on save

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// Loads an existing workbook or creates a new one, applies a red bold font with yellow background to cells A1:B2, and saves the file as XLSX using Aspose.Cells, ensuring all formatting is retained.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // OPTIONAL: Apply some formatting to demonstrate that formatting is retained
            Worksheet sheet = workbook.Worksheets[0];
            Style style = workbook.CreateStyle();
            style.Font.Color = Color.Red;
            style.Font.IsBold = true;
            style.ForegroundColor = Color.Yellow;
            style.Pattern = BackgroundType.Solid;

            // Apply the style to a range (A1:B2)
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1", "B2");
            StyleFlag flag = new StyleFlag
            {
                Font = true,
                CellShading = true
            };
            range.ApplyStyle(style, flag);

            // Save the workbook to an XLSX file, preserving all applied formatting
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
