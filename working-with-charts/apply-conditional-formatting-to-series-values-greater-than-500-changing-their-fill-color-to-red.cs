// Title: How to add a red fill conditional formatting rule for values over 500 in an Aspose.Cells .NET worksheet
// AI Prompts: Write C# code that creates a conditional formatting rule on range A1:A10 to color cells with a solid red fill when their numeric value exceeds 500 using Aspose.Cells. | Generate a reusable Aspose.Cells style with a solid red background and apply it to a greater‑than‑500 condition for a column of data.
// Common Searches: Aspose.Cells C# how to highlight cells greater than 500 with red background | apply conditional formatting to column A based on numeric threshold in Aspose.Cells | set solid fill color for cells exceeding a value using Aspose.Cells .NET API | example of greater than operator in Aspose.Cells conditional formatting
// Tags: add conditional formatting rule Aspose.Cells | red fill style Aspose.Cells | greater-than operator conditional format Aspose.Cells | apply formatting to column A Aspose.Cells | solid background style creation Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;

// Creates a workbook, fills column A with values 0‑900, defines a conditional formatting rule for A1:A10 that applies a solid red fill when the cell value is greater than 500, and saves the file as ConditionalFormatting.xlsx.
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

            // Populate sample data (optional)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i * 100); // Values: 0,100,...,900
            }

            // Add a conditional formatting rule to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            var conditionalFormatting = sheet.ConditionalFormattings[cfIndex];

            // Define the range the rule applies to (A1:A10)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 9,
                EndColumn = 0
            };
            conditionalFormatting.AddArea(area);

            // Add a condition: cell value greater than 500
            int conditionIndex = conditionalFormatting.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "500",
                null); // formula2 not required for this condition

            // Create a style with a solid red fill
            Style redFillStyle = workbook.CreateStyle();
            redFillStyle.ForegroundColor = Color.Red;
            redFillStyle.Pattern = BackgroundType.Solid;

            // Assign the style to the condition
            conditionalFormatting[conditionIndex].Style = redFillStyle;

            // Save the workbook
            string outputPath = "ConditionalFormatting.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
