using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample values
        worksheet.Cells["A1"].PutValue(5);
        worksheet.Cells["A2"].PutValue(10);
        worksheet.Cells["A3"].PutValue(15);

        // Add formulas that reference each other to build a calculation tree
        worksheet.Cells["B1"].Formula = "=A1*2";          // Child of A1
        worksheet.Cells["B2"].Formula = "=A2*2";          // Child of A2
        worksheet.Cells["B3"].Formula = "=A3*2";          // Child of A3
        worksheet.Cells["C1"].Formula = "=B1+B2+B3";      // Parent of B1, B2, B3
        worksheet.Cells["D1"].Formula = "=C1+100";        // Root node

        // Calculate all formulas so that the workbook has evaluated values
        workbook.CalculateFormula();

        // Configure JSON save options to export as a parent‑child hierarchy
        JsonSaveOptions saveOptions = new JsonSaveOptions
        {
            ExportNestedStructure = true,          // Enable nested (tree) JSON output
            AlwaysExportAsJsonObject = true        // Ensure the output is a JSON object even for a single sheet
        };

        // Save the workbook as JSON; the resulting file contains the formula tree
        string outputPath = "formula_tree.json";
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine($"Formula tree exported to: {outputPath}");
    }
}