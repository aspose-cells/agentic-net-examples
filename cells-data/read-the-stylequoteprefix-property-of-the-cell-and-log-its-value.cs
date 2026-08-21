// Title: Read the QuotePrefix flag of a cell with Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, set a leading apostrophe in cell B10, enable the QuotePrefix style, save and reload the file, then read the QuotePrefix property from the cell's style and print the boolean result.
// Keywords: Aspose.Cells QuotePrefix | C# QuotePrefix property | read cell style Aspose.Cells | leading apostrophe Excel | check QuotePrefix flag | Aspose.Cells get style | QuotePrefix example
// Common Searches: Aspose.Cells read QuotePrefix | How to get QuotePrefix flag in C# | Check if cell has leading apostrophe using Aspose.Cells | Retrieve cell style QuotePrefix after loading workbook
// Developer Intent: Retrieve the QuotePrefix boolean flag of a specific cell to determine whether the leading apostrophe formatting is applied.
// Use Cases: Validate that imported data preserves leading apostrophe formatting by inspecting the QuotePrefix flag after loading a workbook. | Apply conditional formatting only when a cell's QuotePrefix property is true. | Log QuotePrefix status of key cells for audit or debugging during spreadsheet processing.
// AI Prompts: Generate C# code with Aspose.Cells that reads the QuotePrefix property of cell B10 after opening a workbook. | Explain how to detect a leading apostrophe in a cell using the QuotePrefix style flag in Aspose.Cells for .NET. | Provide a loop that iterates over a range of cells and outputs each cell's QuotePrefix value.

using System;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixDemo
{
    // Shows how to create a workbook, set a leading apostrophe in cell B10, enable the QuotePrefix style, save and reload the file, then read the QuotePrefix property from the cell's style and print the boolean result.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();                     // create
            Worksheet worksheet = workbook.Worksheets[0];
            Cell cell = worksheet.Cells["B10"];

            // Set a value that starts with a single quote
            cell.PutValue("'12345");

            // Enable the QuotePrefix style for the cell
            Style style = cell.GetStyle();                         // get style
            style.QuotePrefix = true;
            cell.SetStyle(style);                                  // apply style

            // Save the workbook (create a physical file)
            string filePath = "QuotePrefixDemo.xlsx";
            workbook.Save(filePath);                               // save

            // ---------- Load the workbook ----------
            Workbook loadedWorkbook = new Workbook(filePath);       // load
            Cell loadedCell = loadedWorkbook.Worksheets[0].Cells["B10"];

            // Read the QuotePrefix property and log its value
            bool quotePrefix = loadedCell.GetStyle().QuotePrefix; // read property
            Console.WriteLine("QuotePrefix is set: " + quotePrefix);
        }
    }
}
