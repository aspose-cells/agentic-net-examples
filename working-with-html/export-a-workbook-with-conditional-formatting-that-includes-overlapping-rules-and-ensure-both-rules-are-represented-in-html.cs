// Title: Export an Excel workbook with overlapping conditional formatting rules to HTML using Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds two overlapping conditional formatting rules to the same cell range and saves the workbook as an HTML file with Aspose.Cells. | Demonstrate how to keep both conditional formatting styles when converting an Excel workbook to HTML in a .NET application.
// Common Searches: how to export Excel to HTML with multiple conditional formatting rules using Aspose.Cells C# | Aspose.Cells preserve overlapping conditional formats when saving as HTML | C# Aspose.Cells export conditional formatting range A1:A10 to HTML
// Tags: export workbook with overlapping conditional formatting to HTML | Aspose.Cells conditional formatting HTML conversion | C# add multiple conditional format rules | save Excel as HTML preserving conditional styles | conditional formatting range A1:A10 Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Creates a workbook, applies two overlapping conditional formatting rules (greater than 50 and less than 20) to cells A1:A10, and exports the workbook to HTML while retaining both formatting rules using Aspose.Cells for .NET.
class ConditionalFormattingExport
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in range A1:A10
            for (int i = 0; i < 10; i++)
            {
                cells[i, 0].PutValue(i * 10); // Values: 0,10,20,...,90
            }

            // Define the range to which the conditional formats will be applied
            CellArea formatRange = CellArea.CreateCellArea("A1", "A10");

            // First conditional format: highlight cells > 50 with light green background
            int index1 = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc1 = sheet.ConditionalFormattings[index1];
            fcc1.AddArea(formatRange);
            FormatCondition condition1 = fcc1[0];
            condition1.Type = FormatConditionType.CellValue;
            condition1.Operator = OperatorType.GreaterThan;
            condition1.Formula1 = "50";
            condition1.Style = workbook.CreateStyle();
            condition1.Style.ForegroundColor = Color.LightGreen;
            condition1.Style.Pattern = BackgroundType.Solid;

            // Second conditional format: highlight cells < 20 with light coral background
            int index2 = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc2 = sheet.ConditionalFormattings[index2];
            fcc2.AddArea(formatRange);
            FormatCondition condition2 = fcc2[0];
            condition2.Type = FormatConditionType.CellValue;
            condition2.Operator = OperatorType.LessThan; // Correct enum value
            condition2.Formula1 = "20";
            condition2.Style = workbook.CreateStyle();
            condition2.Style.ForegroundColor = Color.LightCoral;
            condition2.Style.Pattern = BackgroundType.Solid;

            // Save the workbook as HTML (conditional formatting is exported by default)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            workbook.Save("ConditionalFormatting.html", htmlOptions);

            Console.WriteLine("Workbook exported to ConditionalFormatting.html with overlapping conditional formats.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
