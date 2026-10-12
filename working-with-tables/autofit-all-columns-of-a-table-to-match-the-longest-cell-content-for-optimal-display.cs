// Title: Auto‑fit every column in an Excel worksheet to the longest cell value with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xlsx file using Aspose.Cells, determines the last used column, calls Worksheet.AutoFitColumn for each column, and saves the updated workbook. | Create a reusable C# method that receives a Worksheet object and automatically adjusts the width of all its columns based on the maximum content length with Aspose.Cells.
// Common Searches: C# Aspose.Cells auto fit column width for entire worksheet | How to adjust Excel column widths to content using Aspose.Cells .NET | Iterate through columns and auto‑size them with Aspose.Cells in C# | Determine last used column and auto‑fit columns in Aspose.Cells workbook | Save workbook after auto‑fitting columns with Aspose.Cells
// Tags: Worksheet.AutoFitColumn C# | auto‑fit columns Aspose.Cells | last used column detection Aspose.Cells | adjust Excel column width .NET | save workbook after column auto‑fit Aspose.Cells

using Aspose.Cells;
using System;

// The example loads 'input.xlsx', finds the last used column, auto‑fits each column to its longest cell using Worksheet.AutoFitColumn, and saves the result as 'output.xlsx' with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (or specify the desired one)
        Worksheet sheet = workbook.Worksheets[0];

        // Determine the last used column in the worksheet
        int lastUsedColumn = sheet.Cells.MaxColumn;

        // Auto‑fit each column from the first to the last used column
        for (int col = 0; col <= lastUsedColumn; col++)
        {
            sheet.AutoFitColumn(col);
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
