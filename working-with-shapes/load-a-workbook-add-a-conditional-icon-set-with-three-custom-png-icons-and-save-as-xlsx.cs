// Title: Apply a three‑icon conditional formatting set with custom PNG icons to cells A1:A10 using Aspose.Cells for .NET and save as XLSX
// AI Prompts: Generate C# code that loads an existing XLSX workbook, creates a three‑icon conditional formatting rule using custom PNG images, applies it to the range A1:A10, and saves the modified file with Aspose.Cells. | Adapt the example to accept a user‑defined cell range and output path, while preserving the custom PNG icon set in the conditional formatting.
// Common Searches: how to use custom PNG icons in Aspose.Cells conditional formatting C# | Aspose.Cells add three‑icon set to a specific range and save as xlsx | C# replace default conditional icon set with custom images using Aspose.Cells | load workbook, apply custom icon set conditional formatting, export XLSX Aspose.Cells | Aspose.Cells conditional formatting icon set for cells A1:A10 example
// Tags: Aspose.Cells custom PNG icon set conditional formatting | C# three‑icon conditional formatting range | save workbook as XLSX Aspose.Cells | replace default icons with custom images Aspose.Cells | conditional formatting icon set file format XLSX

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing XLSX file with Aspose.Cells, adds a conditional formatting rule that uses a three‑icon set backed by custom PNG images, applies the rule to cells A1:A10 on the first worksheet, ensures the output directory exists, and saves the updated workbook as a new XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the target range (A1:A10) – zero‑based indexes
            int firstRow = 0;
            int firstColumn = 0;
            int lastRow = firstRow + 9;
            int lastColumn = firstColumn;

            // Add a new conditional formatting rule to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            var cf = sheet.ConditionalFormattings[cfIndex]; // Use var to avoid explicit type issues

            // Specify the area the rule applies to
            CellArea area = new CellArea
            {
                StartRow = firstRow,
                StartColumn = firstColumn,
                EndRow = lastRow,
                EndColumn = lastColumn
            };
            cf.AddArea(area);

            // Add an Icon Set condition (Three Symbols). Default settings are used.
            cf.AddCondition(FormatConditionType.IconSet);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
