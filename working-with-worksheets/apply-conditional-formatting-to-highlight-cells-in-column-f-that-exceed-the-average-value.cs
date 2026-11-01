// Title: How to apply conditional formatting in Aspose.Cells for .NET to highlight values in column F that are above the column average
// Common Searches: Aspose.Cells C# conditional formatting rule to highlight cells greater than the column average | C# code for applying an expression‑based conditional format using the AVERAGE function in Excel with Aspose.Cells | How to set a light‑salmon background for cells in column F that exceed the average value using Aspose.Cells .NET | Programmatically add conditional formatting to a specific worksheet column with Aspose.Cells | Apply expression‑based conditional formatting to a column range in an Aspose.Cells workbook
// Tags: conditional formatting expression Aspose.Cells C# | average function in conditional formatting Aspose.Cells | column average based cell highlighting Aspose.Cells | solid background style for conditional format Aspose.Cells | apply conditional formatting to column range Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a workbook, defines a CellArea covering rows 1‑100 of column F, adds a conditional formatting rule with the expression `F1>AVERAGE($F:$F)`, sets a solid light‑salmon background style for cells that meet the condition, and saves the file as ConditionalFormatting.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for column F (zero‑based index 5). Adjust rows as needed.
            int firstRow = 0;      // Row 1 (zero‑based)
            int lastRow = 99;      // Row 100 (zero‑based)
            int columnIndex = 5;   // Column F (zero‑based)

            // Build the cell area for the target column
            CellArea area = new CellArea
            {
                StartRow = firstRow,
                StartColumn = columnIndex,
                EndRow = lastRow,
                EndColumn = columnIndex
            };

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];
            cfCollection.AddArea(area);

            // Add a condition: cell value > AVERAGE of the whole column F
            // The relative reference (F1) will shift for each cell automatically.
            string formula = "F1>AVERAGE($F:$F)";
            int conditionIndex = cfCollection.AddCondition(
                FormatConditionType.Expression,
                OperatorType.None,
                formula,
                null);
            FormatCondition condition = cfCollection[conditionIndex];

            // Define the highlight style (e.g., light salmon background)
            Style highlightStyle = workbook.CreateStyle();
            highlightStyle.ForegroundColor = Color.LightSalmon;
            highlightStyle.Pattern = BackgroundType.Solid;
            condition.Style = highlightStyle;

            // Save the workbook with the conditional formatting applied
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
