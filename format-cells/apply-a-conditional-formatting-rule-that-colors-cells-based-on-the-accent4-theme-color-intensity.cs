// Title: How to apply Accent4 theme color conditional formatting to a cell range using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set Accent4 theme color conditional formatting for range A1:A10 | Apply solid fill style with Accent4 color in Aspose.Cells conditional format | Create an always‑true expression conditional format using Aspose.Cells .NET | Use Color.FromArgb to mimic Excel Accent4 in Aspose.Cells conditional formatting | Aspose.Cells example for theme‑based conditional formatting in C#
// Tags: conditional formatting accent4 aspnet | solid fill style conditional format aspnet | expression based conditional format aspnet | apply conditional format to cell range aspnet | xlsx conditional formatting aspnet

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example creates a new workbook, populates cells A1 through A10 with numbers, adds a conditional formatting rule that always evaluates to TRUE, and applies a solid fill style using an Accent4‑like RGB color to the specified range. The workbook is saved as ConditionalFormatting_Accent4.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in A1:A10
            for (int i = 0; i < 10; i++)
            {
                cells[i, 0].PutValue(i);
            }

            // Define the range for conditional formatting (A1:A10)
            CellArea area = CellArea.CreateCellArea("A1", "A10");

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];
            fcc.AddArea(area);

            // Add a condition that always evaluates to true (Expression "TRUE")
            FormatCondition condition = fcc[0];
            condition.Type = FormatConditionType.Expression;
            condition.Formula1 = "TRUE";

            // Configure the style: solid fill with a color similar to Accent4
            Style style = condition.Style;
            style.Pattern = BackgroundType.Solid;                     // Solid fill
            style.ForegroundColor = Color.FromArgb(0, 112, 192);      // Approximate Accent4 color

            // Apply the style back to the condition
            condition.Style = style;

            // Determine output file path
            string outputPath = "ConditionalFormatting_Accent4.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
