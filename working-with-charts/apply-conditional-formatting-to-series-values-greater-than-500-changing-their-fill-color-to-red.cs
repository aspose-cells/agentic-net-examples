// Title: Apply red fill conditional formatting to cells with values greater than 500 using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# conditional formatting cells greater than 500 red fill | How to set a solid red background for values over 500 with Aspose.Cells | C# example of adding a GreaterThan format condition in an Excel workbook using Aspose.Cells | Applying conditional formatting to a specific range (A1:A10) in Aspose.Cells for .NET | Create a style with red fill and attach it to a CellValue condition in Aspose.Cells C#
// Tags: Aspose.Cells conditional formatting style red background | C# format condition GreaterThan operator Aspose.Cells | Excel workbook solid fill style creation Aspose.Cells | Conditional format for column A cells Aspose.Cells | CellValue condition threshold 500 Aspose.Cells | Aspose.Cells CreateStyle set foreground color red

using System;
using System.Drawing;
using Aspose.Cells;

// The sample creates a new workbook, fills cells A1:A10 with values ranging from 400 to 600, defines a conditional formatting range for those cells, adds a CellValue condition that triggers when the value is greater than 500, applies a solid red fill style to the matching cells, and saves the workbook as 'ConditionalFormattingDemo.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data in column A (rows 1 to 10)
            Cells cells = worksheet.Cells;
            for (int i = 0; i < 10; i++)
            {
                // Example values ranging from 400 to 600
                int value = 400 + i * 25;
                cells[i, 0].PutValue(value);
            }

            // Define the range to which the conditional formatting will be applied (A1:A10)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 0
            };

            // Add a new Conditional Formatting collection to the worksheet
            int cfIndex = worksheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = worksheet.ConditionalFormattings[cfIndex];

            // Associate the defined area with this conditional formatting
            cfCollection.AddArea(area);

            // Add a condition of type CellValue (greater than 500)
            int conditionIndex = cfCollection.AddCondition(FormatConditionType.CellValue);
            FormatCondition condition = cfCollection[conditionIndex];

            // Set the operator and the comparison value
            condition.Operator = OperatorType.GreaterThan;
            condition.Formula1 = "500";

            // Create a style with solid red fill
            Style redStyle = workbook.CreateStyle();
            redStyle.ForegroundColor = Color.Red;
            redStyle.Pattern = BackgroundType.Solid;

            // Apply the style to the condition
            condition.Style = redStyle;

            // Save the workbook
            string outputPath = "ConditionalFormattingDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
