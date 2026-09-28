// Title: C# example: Find Excel cells that have a leading apostrophe (QuotePrefix) and rich‑text formatting using Aspose.Cells
// AI Prompts: Write C# code with Aspose.Cells that scans the used range of a worksheet and prints the address of every cell where cell.Style.QuotePrefix is true and cell.IsRichText() returns true. | Create a reusable C# method that returns a List<string> of cell names containing both a leading apostrophe and rich‑text formatting, using the Aspose.Cells .NET API. | Extend the detection routine to also capture each matching cell’s font name, size, and color, and output these details together with the cell address in C#.
// Common Searches: Aspose.Cells C# how to identify cells with QuotePrefix and rich text | list Excel cells that start with an apostrophe and have rich‑text formatting using .NET | detect leading apostrophe in Excel cells with Aspose.Cells API | C# iterate used range and check IsRichText and QuotePrefix in Aspose.Cells
// Tags: quote prefix detection Aspose.Cells C# | rich text cell enumeration Aspose.Cells | find cells with leading apostrophe Excel .NET | Aspose.Cells used range iteration | log cell addresses rich text Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsRichTextApostropheDetection
{
    // The program loads an Excel workbook, iterates over all used cells, checks each cell for a leading apostrophe (QuotePrefix) and rich‑text formatting (IsRichText), and writes the addresses of cells meeting both conditions to the console.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Get the first worksheet (adjust if needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Iterate through all used cells in the worksheet
            for (int row = 0; row <= cells.MaxDataRow; row++)
            {
                for (int col = 0; col <= cells.MaxDataColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Check if the cell starts with a leading apostrophe (QuotePrefix)
                    bool hasLeadingApostrophe = cell.GetStyle().QuotePrefix;

                    // Check if the cell contains rich‑text formatting
                    bool isRichText = cell.IsRichText();

                    // If both conditions are true, log the cell address
                    if (hasLeadingApostrophe && isRichText)
                    {
                        Console.WriteLine($"Cell {cell.Name} (Row {cell.Row}, Column {cell.Column}) contains both a leading apostrophe and rich text.");
                    }
                }
            }

            // Optionally, save the workbook (no changes made, just demonstration)
            workbook.Save("output.xlsx");
        }
    }
}
