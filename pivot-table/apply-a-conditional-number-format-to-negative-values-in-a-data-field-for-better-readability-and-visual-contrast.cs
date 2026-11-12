// Title: Apply a red‑font conditional number format to negative values in an Excel column using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set custom number format for negative numbers in a column | C# conditional formatting rule for values less than zero using Aspose.Cells | How to display negative values in red with a custom format in an Aspose.Cells workbook | Apply conditional formatting to range A1:A10 for negative numbers with Aspose.Cells .NET | Aspose.Cells example red font custom format for cells below zero
// Tags: Aspose.Cells negative number conditional format | C# custom number format red font Aspose.Cells | apply conditional formatting to Excel range Aspose.Cells | FormatConditionType CellValue less than zero Aspose.Cells | Excel workbook conditional number format .NET Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The C# program creates a new workbook, fills cells A1‑A10 with values –5 to 4, defines a conditional formatting rule for that range, adds a CellValue condition for values less than zero, and sets the style to a red font with the custom number format "#,##0;[Red]-#,##0;0" before saving the file as ConditionalNegativeFormat.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Populate sample data in column A (A1:A10) with values from -5 to 4
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i - 5);
            }

            // Define the cell area for conditional formatting (A1:A10)
            var area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 9,
                EndColumn = 0
            };

            // Add a conditional formatting rule and associate the area
            int cfIndex = sheet.ConditionalFormattings.Add();
            var cf = sheet.ConditionalFormattings[cfIndex];
            cf.AddArea(area);

            // Add condition: cell value less than 0
            int conditionIndex = cf.AddCondition(FormatConditionType.CellValue, OperatorType.LessThan, "0", "");

            // Retrieve the condition object and set its style
            var condition = cf[conditionIndex];
            var style = condition.Style;
            style.Font.Color = Color.Red;
            style.Custom = "#,##0;[Red]-#,##0;0";

            // Save the workbook
            string outputPath = "ConditionalNegativeFormat.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
