// Title: How to apply conditional formatting in Aspose.Cells for .NET to highlight cells with values greater than a specified threshold
// AI Prompts: Generate C# code using Aspose.Cells that adds a conditional formatting rule to color cells yellow when their numeric value exceeds 50 in the range A1:E10. | Write a method that creates a solid yellow background style and attaches it to a greater‑than condition for a defined CellArea with Aspose.Cells. | Provide a script that saves an Excel workbook after applying a cell‑value based conditional formatting rule with a custom threshold using Aspose.Cells.
// Common Searches: Aspose.Cells C# conditional formatting cells greater than 50 example | How to set background color for cells exceeding a value using Aspose.Cells .NET | Apply conditional formatting to range A1:E10 in an Aspose.Cells workbook | Saving Excel file after adding conditional formatting with Aspose.Cells C#
// Tags: Aspose.Cells conditional formatting cell value greater than | C# Aspose.Cells apply conditional formatting range | Aspose.Cells set solid background style | Aspose.Cells save workbook after formatting | Excel conditional formatting threshold Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

// The program creates a new workbook, fills cells with sample numeric data, defines the range A1:E10, adds a conditional formatting rule that highlights cells with values greater than 50 using a solid yellow background, and saves the file as ConditionalFormatting.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (optional)
            for (int i = 0; i < 10; i++)
            {
                for (int j = 0; j < 5; j++)
                {
                    sheet.Cells[i, j].PutValue(i * 10 + j);
                }
            }

            // Define the range to apply conditional formatting (A1:E10)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 9,
                EndColumn = 4
            };

            // Add a new conditional formatting rule collection
            int cfIndex = sheet.ConditionalFormattings.Add();
            var cf = sheet.ConditionalFormattings[cfIndex];
            cf.AddArea(area);

            // Add condition: cell value > 50
            int conditionIndex = cf.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
            var condition = cf[conditionIndex];

            // Create a style for highlighting
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.Yellow;
            style.Pattern = BackgroundType.Solid;
            condition.Style = style;

            // Determine output file path
            string outputPath = "ConditionalFormatting.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
