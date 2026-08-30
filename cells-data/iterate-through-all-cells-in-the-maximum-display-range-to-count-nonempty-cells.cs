// Title: How to count non‑empty cells in a worksheet’s MaxDisplayRange using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that retrieves a worksheet's MaxDisplayRange and tallies cells whose Cell.Value is not null. | Show a robust snippet that first verifies the MaxDisplayRange is not null, then iterates through it and prints the number of populated cells. | Provide a complete Aspose.Cells example that adds sample data, performs the non‑empty cell count in the maximum display range, and saves the workbook.
// Common Searches: Aspose.Cells C# count cells with values in MaxDisplayRange | how to handle null MaxDisplayRange when counting cells in Aspose.Cells | iterate over all cells in worksheet maximum display range using Aspose.Cells .NET | C# example for counting non‑empty cells in an Excel workbook with Aspose.Cells
// Tags: filled cell count Aspose.Cells | maxdisplayrange traversal C# | maxdisplayrange existence check Aspose.Cells | cell value presence check Aspose.Cells | workbook save after analysis Aspose.Cells

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The program creates a workbook, inserts sample data, obtains the worksheet's MaxDisplayRange, iterates through each cell counting those with a non‑null Value, outputs the count, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data (you can replace or remove this part)
            cells["A1"].PutValue("Header");
            cells["B2"].PutValue(123);
            cells["C3"].PutValue("Sample text");

            // Obtain the maximum display range (covers data, merged cells, shapes)
            AsposeRange maxDisplayRange = cells.MaxDisplayRange;

            // If the worksheet is empty, MaxDisplayRange will be null
            if (maxDisplayRange == null)
            {
                Console.WriteLine("The worksheet contains no data.");
                return;
            }

            int nonEmptyCellCount = 0;

            // Iterate through each cell in the range and count cells that have a value
            foreach (Cell cell in maxDisplayRange)
            {
                if (cell.Value != null)
                {
                    nonEmptyCellCount++;
                }
            }

            Console.WriteLine($"Non‑empty cells in MaxDisplayRange: {nonEmptyCellCount}");

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("NonEmptyCellCountDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
