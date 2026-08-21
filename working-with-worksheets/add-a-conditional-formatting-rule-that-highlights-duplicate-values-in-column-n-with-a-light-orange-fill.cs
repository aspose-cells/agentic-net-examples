// Title: C# – Highlight Duplicate Values in Column N with Light Orange Fill Using Aspose.Cells Conditional Formatting
// Description: Demonstrates how to create a workbook with Aspose.Cells, define a range for column N, add a DuplicateValues conditional‑formatting rule, apply a light orange background to duplicated cells, and save the file as DuplicateValuesHighlight.xlsx.
// Keywords: Aspose.Cells C# duplicate values | conditional formatting column N | highlight duplicates Excel Aspose | light orange cell fill | Aspose.Cells FormatCondition DuplicateValues | C# Excel conditional formatting example
// Common Searches: Aspose.Cells highlight duplicate values in a column | C# conditional formatting duplicate values Excel | set orange background for duplicate cells Aspose | how to add DuplicateValues rule with Aspose.Cells | conditional formatting range column N C#
// Developer Intent: Add a conditional‑formatting rule that marks duplicate entries in column N with a light orange background using Aspose.Cells for .NET.
// Use Cases: Detect repeated product IDs in an inventory sheet by highlighting duplicates in column N. | Prevent duplicate employee numbers in HR reports with visual cues. | Automatically flag duplicate order numbers in a dynamic dataset where rows are added over time.
// AI Prompts: Generate C# code that uses Aspose.Cells to apply a DuplicateValues conditional format to column N with a custom light orange fill and saves the workbook. | Create a reusable method that accepts a Worksheet, column index, and Color, then adds duplicate‑value highlighting for all rows in that column. | Explain how to determine the last used row in a worksheet and set the conditional‑formatting area dynamically for column N.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsConditionalFormattingDemo
{
    // Demonstrates how to create a workbook with Aspose.Cells, define a range for column N, add a DuplicateValues conditional‑formatting rule, apply a light orange background to duplicated cells, and save the file as DuplicateValuesHighlight.xlsx.
    public class HighlightDuplicateValuesInColumnN
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a new conditional formatting collection
                int cfIndex = worksheet.ConditionalFormattings.Add();
                FormatConditionCollection fcs = worksheet.ConditionalFormattings[cfIndex];

                // Define the range for column N (zero‑based column index 13)
                // Here we assume rows 0‑1000; adjust as needed for your data size
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    EndRow = 1000,
                    StartColumn = 13,
                    EndColumn = 13
                };
                fcs.AddArea(area);

                // Add a duplicate‑values condition
                int conditionIndex = fcs.AddCondition(FormatConditionType.DuplicateValues);
                FormatCondition duplicateCondition = fcs[conditionIndex];

                // Set a light orange background for duplicated cells
                duplicateCondition.Style.BackgroundColor = Color.FromArgb(255, 255, 153); // light orange

                // Save the workbook
                workbook.Save("DuplicateValuesHighlight.xlsx");
                Console.WriteLine("Workbook saved as DuplicateValuesHighlight.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            HighlightDuplicateValuesInColumnN.Run();
        }
    }
}
