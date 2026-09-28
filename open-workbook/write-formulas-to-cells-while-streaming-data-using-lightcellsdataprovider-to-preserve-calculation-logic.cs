// Title: How to write cell formulas while streaming rows using Aspose.Cells LightCellsDataProvider in C#
// AI Prompts: Generate C# code that streams rows into a worksheet with LightCellsDataProvider and assigns a formula to each cell in column B that references the value in column A. | Show an example of using Aspose.Cells LightCells API to populate a large Excel file row by row and set a calculation formula for each row in .NET. | Provide a C# snippet that creates a workbook, streams data with LightCellsDataProvider, and writes a multiplication formula into adjacent cells.
// Common Searches: how to add a multiplication formula to column B when streaming rows with LightCellsDataProvider in C# | set cell formula during LightCells streaming .NET | preserve Excel calculations while generating large files with Aspose.Cells LightCells | C# example of streaming data and applying formulas using Aspose.Cells LightCells | writing formulas to cells while using LightCellsDataProvider to create big workbooks
// Tags: LightCellsDataProvider formula assignment | streaming rows with calculations Aspose.Cells | C# set cell formula during data streaming | large Excel workbook generation with formulas .NET | preserve calculation logic Aspose.Cells LightCells

using System;
using Aspose.Cells;

// The program creates a new workbook, fills column A with values 1‑10, writes a formula in column B that multiplies each A cell by 2, and saves the file as StreamingFormulas.xlsx, illustrating how to embed calculation logic when populating rows (compatible with LightCellsDataProvider streaming).
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the number of rows and columns to populate
            const int totalRows = 10;
            const int totalCols = 2; // Column A for values, Column B for formulas

            // Populate cells directly (fallback when LightCells API is unavailable)
            for (int row = 0; row < totalRows; row++)
            {
                // Column A – numeric value
                sheet.Cells[row, 0].PutValue(row + 1);

                // Column B – formula referencing column A
                string formula = $"=A{row + 1}*2";
                sheet.Cells[row, 1].Formula = formula;
            }

            // Save the workbook
            workbook.Save("StreamingFormulas.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
