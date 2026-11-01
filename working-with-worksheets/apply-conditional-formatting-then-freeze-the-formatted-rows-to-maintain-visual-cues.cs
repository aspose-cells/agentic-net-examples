// Title: Apply conditional formatting to a column and freeze top rows in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# apply conditional formatting to column based on value and freeze first rows | C# code to highlight cells greater than 50 with yellow background using Aspose.Cells | Freeze panes after setting conditional formatting in an Aspose.Cells workbook | Create Excel file with conditional formatting and frozen header rows using Aspose.Cells .NET | Aspose.Cells example for conditional formatting and FreezePanes method in C#
// Tags: conditional formatting column C Aspose.Cells | freeze first rows Aspose.Cells .NET | yellow background style conditional formatting Aspose.Cells | apply cell value greater than 50 rule Aspose.Cells | Aspose.Cells workbook creation with formatting

using System;
using Aspose.Cells;
using System.Drawing;

// The sample creates a workbook, populates ten rows of data, adds a conditional formatting rule that colors cells in column C yellow when their value exceeds 50, freezes the first five rows while keeping columns scrollable, and saves the file as ConditionalFormattingAndFreeze.xlsx.
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

            // Populate sample data (10 rows, 3 columns)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue($"Item {i + 1}");
                sheet.Cells[i, 1].PutValue(i * 10);
                sheet.Cells[i, 2].PutValue(i * 15); // Values for conditional formatting
            }

            // Apply conditional formatting to column C (index 2) for values > 50
            int firstRow = 0;
            int lastRow = 9;
            int col = 2;

            // Define the range to which the formatting will be applied
            CellArea area = new CellArea
            {
                StartRow = firstRow,
                EndRow = lastRow,
                StartColumn = col,
                EndColumn = col
            };

            // Add a new conditional formatting collection
            int cfIndex = sheet.ConditionalFormattings.Add();
            var cf = sheet.ConditionalFormattings[cfIndex];
            cf.AddArea(area);

            // Add a condition: cell value greater than 50
            int conditionIndex = cf.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);

            // Define the style for the condition (yellow background)
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.Yellow;
            style.Pattern = BackgroundType.Solid;
            cf[conditionIndex].Style = style;

            // Freeze the first 5 rows (no columns) while keeping the rest of the sheet scrollable
            int freezeRows = 5;
            int totalRows = sheet.Cells.MaxDataRow + 1;
            int totalColumns = sheet.Cells.MaxDataColumn + 1;
            sheet.FreezePanes(freezeRows, 0, totalRows, totalColumns);

            // Save the workbook
            workbook.Save("ConditionalFormattingAndFreeze.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
