// Title: Add red background conditional formatting for values greater than 50 and export the worksheet to HTML with Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, fills column A with numbers, applies a conditional formatting rule that colors cells red when the value exceeds 50, and saves the workbook as an HTML file while preserving the formatting using Aspose.Cells. | Generate a C# example showing how to define a cell‑value based conditional format (greater than 50) with a solid red background for range A1:A10 and then convert the sheet to HTML via Aspose.Cells.
// Common Searches: Aspose.Cells preserve conditional formatting when converting Excel to HTML in C# | C# example conditional formatting red background export to HTML using Aspose.Cells | how to apply cell value based conditional formatting before saving as HTML with Aspose.Cells .NET | export workbook to HTML with conditional rules applied in Aspose.Cells
// Tags: conditional formatting >50 red background Aspose.Cells | export workbook to HTML preserving formatting Aspose.Cells C# | apply cell‑value conditional style range A1:A10 Aspose.Cells | Aspose.Cells HTML conversion with conditional rules | C# conditional formatting before HTML save Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;

// Demonstrates creating a workbook, populating column A with numeric data, adding a conditional formatting rule that applies a solid red background to cells with values greater than 50, and saving the worksheet as an HTML file while retaining the formatting using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column A (A1:A10)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i * 10); // Values: 0,10,20,...,90
            }

            // Add a conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

            // Define the range to which the formatting will apply (A1:A10)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 0
            };
            cfCollection.AddArea(area);

            // Create a condition: cell value greater than 50
            int conditionIndex = cfCollection.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "50",
                null);
            FormatCondition condition = cfCollection[conditionIndex];

            // Define the style to apply when the condition is met (red background)
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.Red;
            style.Pattern = BackgroundType.Solid;
            condition.Style = style;

            // Save the workbook as HTML (lifecycle: save)
            workbook.Save("ConditionalFormattingOutput.html", SaveFormat.Html);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
