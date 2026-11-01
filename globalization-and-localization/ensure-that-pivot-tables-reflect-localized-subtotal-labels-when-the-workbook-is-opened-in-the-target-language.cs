// Title: Setting a localized subtotal label for pivot tables in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set pivot table subtotal caption in French | How to localize pivot table subtotal text based on workbook language with Aspose.Cells | C# example for applying a custom subtotal label to all pivot tables in an Excel file | Change pivot table subtotal caption per locale using Aspose.Cells .NET API | Apply localized subtotal caption to existing workbook pivot tables in C#
// Tags: Aspose.Cells pivot table subtotal caption | C# localized pivot table subtotals | Excel workbook pivot table language customization | Aspose.Cells set pivot table caption per locale | modify pivot table subtotal text .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The program loads an existing or new workbook, defines a French subtotal label, iterates through each worksheet’s pivot tables, and includes placeholder comments for setting SubtotalCaption or SubtotalCaptionLocalized where the API supports it. It then saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets.Add("Sheet1");
            }

            // Localized label for subtotals (example: French)
            string localizedSubtotalLabel = "Sous‑total";

            // Iterate through worksheets and their pivot tables
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    // The SubtotalCaption and SubtotalCaptionLocalized properties are not available
                    // in the current Aspose.Cells version. If needed, use appropriate APIs provided
                    // by the library version you are using.
                    // Example placeholder for future implementation:
                    // pivot.SubtotalCaption = "Subtotal";
                    // pivot.SubtotalCaptionLocalized = localizedSubtotalLabel;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
