// Title: Highlight duplicate values in column Q with red fill using Aspose.Cells for .NET (C#)
// Description: C# code that loads an Excel workbook, defines a range for column Q, adds a DuplicateValues conditional‑formatting rule, creates a solid red style, applies the style to duplicate cells, and saves the file so any repeated entries in column Q are highlighted in red.
// Keywords: Aspose.Cells | duplicate values | conditional formatting | C# | .NET | column Q | red background | Excel highlight duplicates
// Common Searches: Aspose.Cells highlight duplicate cells | C# conditional formatting duplicate values column Q | set red background for duplicate entries in Excel using Aspose | mark duplicate rows in Excel with Aspose.Cells .NET
// Developer Intent: Automatically detect duplicate entries in column Q of an Excel worksheet and emphasize them with a red background using Aspose.Cells for .NET.
// Use Cases: Identify repeated product IDs in a catalog to prevent publishing errors. | Spot duplicate email addresses in a mailing list before a campaign launch. | Flag duplicate invoice numbers in a financial report to avoid processing mistakes.
// AI Prompts: Generate C# code with Aspose.Cells that applies a DuplicateValues conditional‑formatting rule to column Q and uses a solid red fill. | Explain how to change the target column index and the highlight color in the duplicate‑highlighting example. | Provide a step‑by‑step guide to add additional conditional‑formatting rules (e.g., for blanks or unique values) alongside the duplicate‑value rule in the same worksheet.

using System;
using System.Drawing;
using Aspose.Cells;

// C# code that loads an Excel workbook, defines a range for column Q, adds a DuplicateValues conditional‑formatting rule, creates a solid red style, applies the style to duplicate cells, and saves the file so any repeated entries in column Q are highlighted in red.
class DuplicateHighlighter
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Determine the last row with data in column Q (index 16)
        int lastRow = cells.MaxDataRow;

        // Define the range for column Q (from row 0 to lastRow)
        CellArea range = new CellArea
        {
            StartRow = 0,
            EndRow = lastRow,
            StartColumn = 16,   // Column Q (0‑based index)
            EndColumn = 16
        };

        // Add a new conditional formatting collection
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

        // Apply the range to the conditional formatting
        fcs.AddArea(range);

        // Add a condition that highlights duplicate values
        int conditionIdx = fcs.AddCondition(FormatConditionType.DuplicateValues);
        FormatCondition condition = fcs[conditionIdx];

        // Create a style with red background
        Style redStyle = workbook.CreateStyle();
        redStyle.ForegroundColor = Color.Red;
        redStyle.Pattern = BackgroundType.Solid;

        // Assign the style to the condition
        condition.Style = redStyle;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
