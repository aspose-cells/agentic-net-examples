// Title: C# – Highlight Cells Above Average in Column F Using Aspose.Cells Conditional Formatting
// Description: Sample C# code that creates a workbook, fills rows 0‑19 of column F with numeric data, adds an AboveAverage conditional formatting rule, applies a yellow background to values greater than the column average, auto‑fits columns, and saves the file as an XLSX document.
// Keywords: Aspose.Cells | C# | conditional formatting | AboveAverage | highlight cells | column F | Excel average | cell background color | sample code | GitHub example
// Common Searches: Aspose.Cells highlight cells above average C# | Conditional formatting column F Aspose.Cells .NET | How to use AboveAverage format condition in Aspose.Cells | C# code to color cells above average in Excel | Aspose.Cells example for average‑based formatting
// Developer Intent: Apply a conditional formatting rule that colors cells in column F whose values exceed the column’s average.
// Use Cases: Mark sales entries that are higher than the average sales figure in a monthly report. | Highlight student scores that surpass the class average for quick performance review. | Identify inventory items with quantities above the average stock level to spot overstock.
// AI Prompts: Generate C# code with Aspose.Cells that applies an AboveAverage conditional format to a specified column and sets a custom fill color. | Explain how to modify the rule to also include cells equal to the average or to use a different highlight color. | Provide step‑by‑step instructions for retrieving a column’s average value and applying conditional formatting based on that value with Aspose.Cells.

using System;
using Aspose.Cells;
using System.Drawing;

// Sample C# code that creates a workbook, fills rows 0‑19 of column F with numeric data, adds an AboveAverage conditional formatting rule, applies a yellow background to values greater than the column average, auto‑fits columns, and saves the file as an XLSX document.
class HighlightAboveAverageInColumnF
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data in column F (zero‑based index 5)
        for (int row = 0; row < 20; row++)
        {
            worksheet.Cells[row, 5].PutValue(row * 10 + 5);
        }

        // Add a conditional formatting collection to the worksheet
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = worksheet.ConditionalFormattings[cfIndex];

        // Define the range: column F, rows 0‑19
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 19,
            StartColumn = 5,
            EndColumn = 5
        };
        fcc.AddArea(area);

        // Add an AboveAverage condition
        int conditionIndex = fcc.AddCondition(FormatConditionType.AboveAverage);
        FormatCondition fc = fcc[conditionIndex];

        // Configure the condition to highlight values above the average
        fc.AboveAverage.IsAboveAverage = true;      // true = above average
        fc.AboveAverage.IsEqualAverage = false;    // exclude cells equal to the average
        fc.Style.BackgroundColor = Color.Yellow;   // highlight color

        // Optional: auto‑fit columns for better appearance
        worksheet.AutoFitColumns();

        // Save the workbook
        workbook.Save("ColumnF_AboveAverageHighlight.xlsx");
    }
}
