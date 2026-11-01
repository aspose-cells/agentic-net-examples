// Title: Apply conditional formatting in Aspose.Cells for .NET to highlight column F cells with values greater than 1000
// Common Searches: Aspose.Cells C# conditional formatting highlight cells greater than 1000 in column F | Create yellow fill rule for values above 1000 using Aspose.Cells .NET | Add cell value > 1000 conditional format to range F1:F100 with Aspose.Cells | Set background color for high numeric values in column F using Aspose.Cells API | Conditional formatting example for numeric threshold in Excel with Aspose.Cells C#
// Tags: conditional formatting column F Aspose.Cells | highlight cells over 1000 C# | yellow fill style Aspose.Cells | format condition type CellValue operator GreaterThan | apply conditional formatting Excel .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace ConditionalFormattingExample
{
    // The example creates a workbook, defines the range F1:F100, adds a conditional formatting rule that flags cells with values greater than 1000, applies a solid yellow background style to those cells, and saves the file as ConditionalFormatting.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define the range for column F (F1:F100). Row and column indices are zero‑based.
                CellArea range = CellArea.CreateCellArea(0, 5, 99, 5);

                // Add a conditional formatting collection to the worksheet.
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];
                cfCollection.AddArea(range);

                // Add a condition: cell value > 1000.
                int conditionIndex = cfCollection.AddCondition(
                    FormatConditionType.CellValue,
                    OperatorType.GreaterThan,
                    "1000",
                    ""); // second formula not required for this operator

                FormatCondition condition = cfCollection[conditionIndex];

                // Define the style to apply when the condition is met (yellow fill).
                Style highlightStyle = workbook.CreateStyle();
                highlightStyle.ForegroundColor = Color.Yellow;
                highlightStyle.Pattern = BackgroundType.Solid;
                condition.Style = highlightStyle;

                // Save the workbook.
                string outputPath = "ConditionalFormatting.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
