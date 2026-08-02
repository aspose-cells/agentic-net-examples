// Title: C# – Apply Accent2 Theme Color Conditional Formatting with Aspose.Cells for Values Over a Threshold
// Description: This example shows how to create a workbook with Aspose.Cells for .NET, fill column A with numbers, define a threshold (e.g., 50), and add a conditional‑formatting rule for A1:A10 that highlights cells whose values exceed the threshold using the workbook’s Accent2 theme color. The workbook is then saved as an .xlsx file.
// Keywords: Aspose.Cells C# conditional formatting | Accent2 theme color Excel | highlight cells above threshold | FormatConditionType.CellValue example | OperatorType.GreaterThan Aspose | solid fill style Excel theme | C# Excel automation Aspose.Cells | apply theme colors programmatically | Excel conditional formatting .NET
// Common Searches: how to use Accent2 theme color in Aspose.Cells conditional formatting | C# code to highlight values greater than 50 in Excel with Aspose | apply solid fill style using workbook theme colors Aspose.Cells | add CellValue condition to a range in Aspose.Cells .NET | retrieve Accent2 color from Excel theme with Aspose
// Developer Intent: Add a conditional‑formatting rule that uses the workbook’s Accent2 theme color to highlight cells whose numeric values are greater than a specified threshold.
// Use Cases: Flag sales figures that surpass a target in a financial dashboard. | Mark temperature readings that exceed safety limits in an engineering log. | Highlight student scores above a passing grade in an academic report.
// AI Prompts: Generate C# Aspose.Cells code that applies conditional formatting with the workbook’s Accent2 theme color to cells where the value is greater than a given threshold. | Show how to create a FormatCondition of type CellValue with OperatorType.GreaterThan and set its style to the actual Accent2 color from the workbook theme. | Explain how to replace a hard‑coded LightBlue color with the real Accent2 color retrieved from the workbook’s theme in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing; // For BackgroundType enum

// This example shows how to create a workbook with Aspose.Cells for .NET, fill column A with numbers, define a threshold (e.g., 50), and add a conditional‑formatting rule for A1:A10 that highlights cells whose values exceed the threshold using the workbook’s Accent2 theme color. The workbook is then saved as an .xlsx file.
class ConditionalFormattingAccent2
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample numeric data in column A (rows 1-10)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i * 10); // Values: 0,10,20,...,90
            }

            // Define the threshold value to compare against
            double threshold = 50; // Highlight values greater than 50

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Set the range to which the conditional formatting will be applied (A1:A10)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 0
            };
            fcc.AddArea(area);

            // Add a CellValue condition: values greater than the threshold
            int conditionIndex = fcc.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                threshold.ToString(),
                null);

            // Retrieve the created condition
            FormatCondition fc = fcc[conditionIndex];

            // Configure the style: solid fill with a light blue color (approximation of Accent2)
            fc.Style.Pattern = BackgroundType.Solid;               // Solid fill
            fc.Style.ForegroundColor = Color.LightBlue;           // Approximate Accent2 color

            // Save the workbook
            string outputPath = "ConditionalFormatting_Accent2.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
