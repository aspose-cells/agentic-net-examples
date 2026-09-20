// Title: How to add conditional formatting that highlights duplicate values in column N with a light orange fill using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to apply a duplicate‑value conditional formatting rule to column N and set a light orange background. | Write a program that creates an Excel workbook, defines column N, adds a duplicate values condition with a solid orange fill, and saves the file. | Update an existing Aspose.Cells worksheet to highlight duplicate entries in column N using a light orange fill style.
// Common Searches: aspnet how to highlight duplicate entries in column N using Aspose.Cells | C# Aspose.Cells conditional formatting duplicate values orange fill | programmatically set duplicate value rule for column N in Excel with Aspose.Cells .NET | apply light orange background to duplicate cells in a specific column using Aspose.Cells
// Tags: Aspose.Cells duplicate values conditional formatting | column N conditional formatting Aspose.Cells | light orange fill style Aspose.Cells | C# Excel conditional formatting Aspose.Cells | apply conditional formatting to specific column .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// Creates a new workbook, defines column N as the target range, adds a duplicate‑value conditional formatting rule with a light orange solid fill, and saves the result as HighlightedDuplicates.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range for column N (zero‑based column index 13)
            CellArea area = new CellArea
            {
                StartRow = 0,          // Row 1 (zero‑based)
                EndRow = 1048575,      // Last possible row in Excel
                StartColumn = 13,      // Column N
                EndColumn = 13
            };

            // Add a new ConditionalFormatting object to the worksheet
            int cfIndex = worksheet.ConditionalFormattings.Add();
            var cf = worksheet.ConditionalFormattings[cfIndex]; // ConditionalFormatting instance

            // Associate the defined area with the ConditionalFormatting object
            cf.AddArea(area);

            // Add a condition that highlights duplicate values
            int conditionIndex = cf.AddCondition(FormatConditionType.DuplicateValues);
            var condition = cf[conditionIndex]; // FormatCondition instance

            // Create a style with a light orange solid fill
            Style style = condition.Style;
            style.ForegroundColor = Color.FromArgb(255, 255, 200, 0); // Light orange
            style.Pattern = BackgroundType.Solid;
            condition.Style = style;

            // Define output file path
            string outputPath = "HighlightedDuplicates.xlsx";

            // Ensure the directory for the output file exists (if any)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
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
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
