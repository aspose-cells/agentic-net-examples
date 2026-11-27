// Title: Apply conditional formatting in Aspose.Cells for .NET (C#) to highlight column cells that exceed a numeric threshold
// Common Searches: Aspose.Cells C# conditional formatting highlight cells greater than 100 | Add greater‑than conditional format to column B using Aspose.Cells .NET | C# code to apply yellow background to Excel cells above a threshold with Aspose | Create conditional formatting rule for numeric values in an Aspose.Cells workbook | Set conditional formatting range B1:B10 in Aspose.Cells C# example | How to use FormatConditionCollection to color cells in Aspose.Cells
// Tags: conditional formatting with Aspose.Cells C# | highlight cells exceeding numeric threshold Excel | FormatConditionCollection add greater than condition | apply style to conditional format Aspose.Cells | CellArea range B1:B10 Aspose.Cells | Excel workbook conditional formatting .NET

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a new workbook, fills column B (B1:B10) with incremental numbers, defines a threshold of 100, builds a yellow solid‑fill style, adds a conditional formatting rule to the B column that colors cells with values greater than the threshold, and saves the workbook as ConditionalFormattingDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample numeric data in column B (index 1)
            for (int i = 0; i < 10; i++)
            {
                // Example values: 0, 30, 60, ..., 270
                sheet.Cells[i, 1].PutValue(i * 30);
            }

            // Define the numeric threshold
            double threshold = 100;

            // Create a style that will be applied to cells exceeding the threshold
            Style highlightStyle = workbook.CreateStyle();
            highlightStyle.ForegroundColor = Color.Yellow;
            highlightStyle.Pattern = BackgroundType.Solid;

            // Define the range to which the conditional formatting will be applied (B1:B10)
            int firstRow = 0;   // Row 1 (zero‑based index)
            int lastRow = 9;    // Row 10
            int column = 1;     // Column B (zero‑based index)

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

            // Specify the area for the conditional formatting
            CellArea area = new CellArea
            {
                StartRow = firstRow,
                StartColumn = column,
                EndRow = lastRow,
                EndColumn = column
            };
            cfCollection.AddArea(area);

            // Add a condition: cell value greater than the threshold
            // The AddCondition overload requires a second formula argument; pass empty string when not needed
            int conditionIndex = cfCollection.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                threshold.ToString(),
                string.Empty);

            // Apply the highlight style to the condition
            cfCollection[conditionIndex].Style = highlightStyle;

            // Save the workbook (lifecycle rule)
            workbook.Save("ConditionalFormattingDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
