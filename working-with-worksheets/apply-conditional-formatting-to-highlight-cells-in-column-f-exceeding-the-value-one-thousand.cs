// Title: How to use Aspose.Cells for .NET to apply conditional formatting that highlights column F cells with values over 1000
// AI Prompts: Generate C# code with Aspose.Cells to create a conditional formatting rule that fills cells in column F with yellow when the numeric value exceeds 1000. | Write a method that adds a CellValue > 1000 condition to the range F1:F1000 in an existing workbook using the Aspose.Cells .NET API. | Provide a step‑by‑step script to define a CellArea for column F, attach a greater‑than‑1000 format condition, set a solid yellow style, and save the workbook.
// Common Searches: Aspose.Cells C# example to highlight column F cells with values above 1000 | How to add a greater‑than‑1000 conditional format to a specific column using Aspose.Cells | Programmatic way to set yellow fill for cells exceeding 1000 in an Excel workbook with Aspose.Cells .NET | Create a CellArea for column F and apply numeric threshold formatting in C# Aspose.Cells
// Tags: Aspose.Cells conditional formatting column F | C# greater-than numeric condition Aspose.Cells | highlight cells over 1000 Aspose.Cells | apply solid yellow style Aspose.Cells | Excel conditional format rule Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.Drawing;

// Creates a new workbook, defines the range F1:F1000, adds a conditional formatting rule that applies a solid yellow fill to cells whose value is greater than 1000, and saves the file as ConditionalFormatting.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for column F (F1:F1000)
            CellArea area = new CellArea
            {
                StartRow = 0,      // Row 1 (0‑based index)
                EndRow = 999,      // Row 1000
                StartColumn = 5,   // Column F (0‑based index)
                EndColumn = 5
            };

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];
            cfCollection.AddArea(area);

            // Add a condition: cell value greater than 1000
            int conditionIndex = cfCollection.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "1000",
                null);

            // Retrieve the created condition
            FormatCondition condition = cfCollection[conditionIndex];

            // Define the style to apply when the condition is met (e.g., yellow fill)
            Style highlightStyle = workbook.CreateStyle();
            highlightStyle.ForegroundColor = Color.Yellow;
            highlightStyle.Pattern = BackgroundType.Solid;
            condition.Style = highlightStyle;

            // Save the workbook (lifecycle save rule)
            workbook.Save("ConditionalFormatting.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
