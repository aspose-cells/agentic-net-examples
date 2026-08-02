// Title: Aspose.Cells for .NET: Conditional Formatting with Freeze Panes in C#
// Description: C# sample that creates a workbook, populates column A with numeric data, adds a conditional‑format rule to highlight values greater than 50 in light‑green bold cells (rows 1‑5), freezes the first five rows at cell A6, and saves the result as ConditionalFormattingAndFreeze.xlsx.
// Keywords: Aspose.Cells | C# | .NET | conditional formatting | freeze panes | freeze rows | highlight cells | Excel automation | sample code | Workbook | Worksheet | FormatCondition | CellArea | GitHub example
// Common Searches: Aspose.Cells conditional formatting C# example | how to freeze top rows with Aspose.Cells .NET | highlight cells greater than 50 using Aspose.Cells | freeze panes at A6 Aspose.Cells C# | apply conditional format then freeze panes Aspose
// Developer Intent: Create an Excel workbook, apply a value‑based conditional format, and lock the header rows so they stay visible while scrolling, using Aspose.Cells for .NET.
// Use Cases: Sales dashboards where high‑volume rows are colored and header rows remain fixed for quick comparison. | Financial statements that emphasize expenses above a threshold while keeping summary rows in view. | Exported data reports that automatically format critical values and preserve navigation context with frozen panes.
// AI Prompts: Show how to change the conditional formatting to use a red background for values less than 20. | Provide code to freeze both rows and columns at cell C10 while retaining existing formatting rules. | Explain how to add multiple conditional formatting rules to separate ranges before applying freeze panes.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsConditionalFormattingAndFreeze
{
    // C# sample that creates a workbook, populates column A with numeric data, adds a conditional‑format rule to highlight values greater than 50 in light‑green bold cells (rows 1‑5), freezes the first five rows at cell A6, and saves the result as ConditionalFormattingAndFreeze.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in column A (rows 0-9)
            for (int i = 0; i < 10; i++)
            {
                cells[i, 0].PutValue(i * 10); // Values: 0,10,20,...,90
            }

            // ---------- Conditional Formatting ----------
            // Add a new conditional formatting rule collection
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Define the range to which the formatting will be applied (rows 0-4, column A)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 4,
                StartColumn = 0,
                EndColumn = 0
            };
            fcc.AddArea(area);

            // Add a condition: highlight cells with value > 50
            int condIndex = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
            FormatCondition condition = fcc[condIndex];
            condition.Style.BackgroundColor = Color.LightGreen;
            condition.Style.Font.IsBold = true;

            // ---------- Freeze Panes ----------
            // Freeze the top 5 rows so that the conditional formatting remains visible while scrolling
            // Freeze at cell A6 (row index 5) with 5 frozen rows and 0 frozen columns
            sheet.FreezePanes("A6", 5, 0);

            // Save the workbook
            workbook.Save("ConditionalFormattingAndFreeze.xlsx");
        }
    }
}
