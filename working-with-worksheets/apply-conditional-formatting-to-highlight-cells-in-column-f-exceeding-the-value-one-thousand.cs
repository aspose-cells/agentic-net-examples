// Title: How to use Aspose.Cells for .NET (C#) to apply conditional formatting that highlights cells in column F with values greater than 1000
// Common Searches: Aspose.Cells C# conditional formatting column F greater than 1000 | C# set yellow background for Excel cells over 1000 using Aspose.Cells | Apply cell value based conditional formatting in Aspose.Cells .NET | Highlight high numeric values in a specific column with Aspose.Cells API | Create conditional formatting rule for column F in an Excel workbook using C#
// Tags: Aspose.Cells conditional formatting column F | cell value greater than 1000 style Aspose.Cells | C# Excel conditional formatting API | apply background color based on numeric threshold Aspose.Cells | format condition CellValue .NET

using System;
using Aspose.Cells;
using System.Drawing;

// The example creates or loads a workbook, populates column F with numeric data, defines a CellArea covering the desired rows, adds a conditional formatting rule that triggers when a cell value exceeds 1000, applies a yellow solid background style to those cells, and saves the workbook as an XLSX file using the Aspose.Cells for .NET API.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            // If you have an existing file, replace the line below with:
            // Workbook workbook = new Workbook("input.xlsx");
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate column F (index 5) with sample data for demonstration
            // In a real scenario you would already have data in the column.
            Cells cells = sheet.Cells;
            for (int row = 0; row < 20; row++)
            {
                // Example values: some below and some above 1000
                double value = (row % 5 == 0) ? 1500 + row * 10 : 500 + row * 20;
                cells[row, 5].PutValue(value); // Column F = index 5
            }

            // Define the range that covers the entire column F (e.g., rows 1‑100)
            // Adjust the end row as needed for your data size.
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 99,   // rows 1‑100 (0‑based index)
                StartColumn = 5,
                EndColumn = 5
            };

            // Add a conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];
            fcs.AddArea(area);

            // Add a condition: Cell value greater than 1000
            int conditionIndex = fcs.AddCondition(FormatConditionType.CellValue);
            FormatCondition condition = fcs[conditionIndex];
            condition.Operator = OperatorType.GreaterThan;
            condition.Formula1 = "1000";

            // Define the style to apply when the condition is met
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.Yellow;
            style.Pattern = BackgroundType.Solid;
            condition.Style = style;

            // Save the workbook
            string outputPath = "ConditionalFormattingDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while applying conditional formatting:");
            Console.WriteLine(ex.Message);
        }
    }
}
