// Title: Apply red fill conditional formatting to cells with values over 500 using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# conditional formatting cells greater than 500 red background | How to set a solid red fill style for a conditional format rule in Aspose.Cells | C# example adding FormatCondition for values above 500 in an Excel workbook with Aspose.Cells | Apply conditional formatting to range A1:A10 using Aspose.Cells .NET | Save workbook with conditional formatting rule using Aspose.Cells for .NET
// Tags: conditional formatting cellvalue greaterthan Aspose.Cells C# | red fill style background Aspose.Cells | formatconditiontype cellvalue operator greaterthan | apply conditional formatting to range Aspose.Cells | save workbook with conditional formatting Aspose.Cells | create style solid red Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

namespace AsposeCellsExample
{
    // The example creates a workbook, populates column A with incremental values, adds a conditional formatting rule for cells A1:A10 that are greater than 500, applies a solid red fill style to those cells, and saves the workbook as ConditionalFormatting.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate sample data in column A (A1:A10)
                for (int row = 0; row < 10; row++)
                {
                    sheet.Cells[row, 0].PutValue(row * 100 + 50);
                }

                // Add a new conditional formatting rule to the worksheet
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection conditions = sheet.ConditionalFormattings[cfIndex];

                // Define the range the rule applies to (A1:A10)
                CellArea area = CellArea.CreateCellArea("A1", "A10");
                conditions.AddArea(area);

                // Add a condition: cell value greater than 500
                int conditionIndex = conditions.AddCondition(
                    FormatConditionType.CellValue,
                    OperatorType.GreaterThan,
                    "500",
                    null);

                // Retrieve the created condition
                FormatCondition condition = conditions[conditionIndex];

                // Create a style with a solid red fill
                Style redFillStyle = workbook.CreateStyle();
                redFillStyle.ForegroundColor = Color.Red;
                redFillStyle.Pattern = BackgroundType.Solid;

                // Assign the style to the condition
                condition.Style = redFillStyle;

                // Define output path (saved in the current directory)
                string outputPath = "ConditionalFormatting.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
