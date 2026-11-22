// Title: Add a French‑localized total row to an Excel worksheet using Aspose.Cells for .NET
// Common Searches: how to set workbook culture to French in Aspose.Cells C# and calculate sum | Aspose.Cells add total row with SUM formula after setting GlobalizationSettings | C# generate Excel file with French number formatting using Aspose.Cells | save Excel workbook with localized labels and total row using Aspose.Cells .NET | calculate totals in French locale with Aspose.Cells workbook Settings.CultureInfo
// Tags: Aspose.Cells workbook culture fr-FR | C# add total row using SUM formula | French number formatting in Excel with Aspose.Cells | recalculate formulas after culture change Aspose.Cells | save Excel file SubtotalLocalized.xlsx Aspose.Cells

using Aspose.Cells;
using System;
using System.Globalization;
using System.IO;

// The example creates a new workbook, fills it with sample data, sets the workbook's CultureInfo to French (fr-FR) for localized number formatting, inserts a total row with a SUM formula, recalculates formulas, ensures the output directory exists, and saves the file as SubtotalLocalized.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Food");
            sheet.Cells["B2"].PutValue(10.5);
            sheet.Cells["A3"].PutValue("Food");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Transport");
            sheet.Cells["B4"].PutValue(15);
            sheet.Cells["A5"].PutValue("Transport");
            sheet.Cells["B5"].PutValue(25);

            // Set workbook culture to French (fr-FR) for localized formatting
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Add a simple total row at the bottom (since Subtotal API may not be available)
            int totalRowIndex = sheet.Cells.MaxDataRow + 1;
            sheet.Cells[totalRowIndex, 0].PutValue("Total");
            sheet.Cells[totalRowIndex, 1].Formula = $"SUM(B2:B{sheet.Cells.MaxDataRow + 1})";

            // Recalculate formulas to ensure the total is computed
            workbook.CalculateFormula();

            // Save the workbook
            string outputPath = "SubtotalLocalized.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
