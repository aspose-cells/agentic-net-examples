// Title: Highlight duplicate values in column Q with a red background using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to add a DuplicateValues conditional format to column Q and set a solid red fill for the matching cells. | Generate a .NET program that loads an Excel workbook, applies a red background style to any duplicate entries in column Q via conditional formatting, and saves the file.
// Common Searches: Aspose.Cells C# highlight duplicate cells in column Q with red fill | How to apply conditional formatting for duplicate values in a specific column using Aspose.Cells .NET | Set red background for duplicate entries in Excel column Q programmatically with Aspose.Cells | C# example for duplicate value detection and styling in column Q using Aspose.Cells
// Tags: Aspose.Cells duplicate values conditional formatting | red fill style for duplicate cells .NET | column Q duplicate detection Aspose.Cells | conditional formatting red background Excel .NET | highlight duplicate entries column Q C#

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsDuplicateHighlight
{
    // The example loads or creates a workbook, defines a conditional formatting range for column Q, adds a DuplicateValues condition, applies a solid red fill style to the duplicates, and saves the modified file as output.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook or load an existing one
            // Replace "input.xlsx" with your source file if needed
            Workbook workbook = new Workbook(); // new workbook
            // If you have an existing file, use:
            // Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine the last row that contains data in column Q (index 16)
            int lastRow = worksheet.Cells.MaxDataRow;

            // Define the range that covers column Q from the first row to the last data row
            CellArea duplicateRange = new CellArea
            {
                StartRow = 0,
                EndRow = lastRow,
                StartColumn = 16,   // Column Q (0‑based index)
                EndColumn = 16
            };

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = worksheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = worksheet.ConditionalFormattings[cfIndex];

            // Apply the range to the conditional formatting
            fcs.AddArea(duplicateRange);

            // Add a condition that highlights duplicate values
            int conditionIndex = fcs.AddCondition(FormatConditionType.DuplicateValues);
            FormatCondition duplicateCondition = fcs[conditionIndex];

            // Create a style with a red background
            Style redStyle = workbook.CreateStyle();
            redStyle.ForegroundColor = Color.Red;
            redStyle.Pattern = BackgroundType.Solid;

            // Assign the style to the condition
            duplicateCondition.Style = redStyle;

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}
