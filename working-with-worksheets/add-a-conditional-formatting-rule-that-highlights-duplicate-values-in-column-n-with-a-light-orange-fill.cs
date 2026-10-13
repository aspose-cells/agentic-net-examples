// Title: Apply a duplicate‑values conditional format with a light orange fill to column N using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook, creates a DuplicateValues conditional formatting rule for column N, sets a solid light orange background style, and saves the file with Aspose.Cells. | Write a C# snippet using Aspose.Cells to highlight duplicate entries in column N by applying a light orange fill through conditional formatting.
// Common Searches: Aspose.Cells C# highlight duplicate cells in a specific column with orange background | How to add duplicate values conditional formatting to column N in a .NET Excel workbook | C# code sample for applying light orange fill to duplicate values using Aspose.Cells | Conditional formatting duplicate values column N Aspose.Cells example
// Tags: duplicate values conditional formatting Aspose.Cells .NET | apply solid fill style Aspose.Cells | conditional format column N Excel C# | highlight duplicate entries Aspose.Cells | light orange background style Aspose.Cells

using Aspose.Cells;
using System.Drawing;

// // Loads input.xlsx, adds a DuplicateValues conditional formatting rule to column N (index 13) with a solid light orange fill, and saves the workbook as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the range for column N (zero‑based column index 13)
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = sheet.Cells.MaxDataRow,
            StartColumn = 13,
            EndColumn = 13
        };

        // Add a new conditional formatting collection
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];
        cfCollection.AddArea(area);

        // Add a duplicate‑values condition
        int conditionIndex = cfCollection.AddCondition(
            FormatConditionType.DuplicateValues,
            OperatorType.None,
            string.Empty,
            string.Empty);

        // Create a style with a light orange fill
        Style orangeStyle = workbook.CreateStyle();
        orangeStyle.Pattern = BackgroundType.Solid;
        orangeStyle.ForegroundColor = Color.FromArgb(255, 255, 200, 0); // light orange

        // Apply the style to the conditional format
        cfCollection[conditionIndex].Style = orangeStyle;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
