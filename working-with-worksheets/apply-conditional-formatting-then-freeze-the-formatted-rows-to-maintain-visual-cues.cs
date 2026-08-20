// Title: Conditional Formatting and Freeze Panes with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills column A with values 0‑90, adds a conditional formatting rule that colors cells yellow when the value is ≥ 50, freezes the first ten rows and the first column, and saves the file as an XLSX document.
// Keywords: Aspose.Cells conditional formatting C# | freeze panes Aspose.Cells .NET | highlight cells >= 50 Excel | Workbook FreezePanes method | FormatConditionCollection example | CellArea range Aspose.Cells | C# Excel automation | Aspose.Cells API tutorial
// Common Searches: Aspose.Cells C# conditional formatting >= 50 | How to freeze top rows with Aspose.Cells | Freeze panes after applying conditional formatting in .NET | C# example for conditional formatting and FreezePanes | Aspose.Cells highlight cells and keep them visible
// Developer Intent: Generate an XLSX file, apply a yellow background to cells with values ≥ 50, and freeze the first ten rows plus the first column so the highlighted rows stay in view.
// Use Cases: Mark high‑value rows in a sales report while keeping them visible during scrolling. | Create a financial dashboard where rows meeting a threshold are colored and locked for quick comparison. | Build a template that automatically emphasizes key rows with conditional formatting and freeze panes for end‑users.
// AI Prompts: Show how to change the conditional formatting to a red background for values > 80 and freeze only the top row using Aspose.Cells. | Provide a C# snippet that adds three conditional formatting rules (green < 30, yellow 30‑60, red > 60) and freezes rows and columns based on a dynamic range. | Explain how to apply conditional formatting to multiple columns and then freeze the header row and first column in Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// Creates a workbook, fills column A with values 0‑90, adds a conditional formatting rule that colors cells yellow when the value is ≥ 50, freezes the first ten rows and the first column, and saves the file as an XLSX document.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample numeric data in column A (rows 0‑9)
            for (int i = 0; i < 10; i++)
            {
                cells[i, 0].PutValue(i * 10); // 0, 10, 20, ... 90
            }

            // ---------- Conditional Formatting ----------
            // Add a new conditional formatting rule collection
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Define the range A1:A10 (rows 0‑9, column 0)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 0
            };
            fcc.AddArea(area);

            // Add a condition: cells with value >= 50 get a yellow background
            // Use GreaterThan operator with value 49 to emulate >= 50 (compatible with older API versions)
            int condIdx = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "49", null);
            FormatCondition fc = fcc[condIdx];
            fc.Style.BackgroundColor = Color.Yellow;

            // ---------- Freeze Panes ----------
            // Freeze the first 10 rows and the first column so the formatted rows stay visible
            sheet.FreezePanes(10, 1, 10, 1);

            // Save the workbook
            string outputPath = "ConditionalFormattingAndFreeze.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
