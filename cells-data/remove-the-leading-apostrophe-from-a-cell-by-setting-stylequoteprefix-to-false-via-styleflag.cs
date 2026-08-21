// Title: Remove Leading Apostrophe from an Excel Cell using Aspose.Cells C# (Style.QuotePrefix & StyleFlag)
// Description: Demonstrates how to delete the leading single‑quote in a cell (e.g., B2) by creating a Style with QuotePrefix set to false, enabling the property with a StyleFlag, applying it to the cell, and saving the workbook as RemovedApostrophe.xlsx.
// Keywords: Aspose.Cells remove apostrophe | Style.QuotePrefix false C# | StyleFlag QuotePrefix | Excel leading single quote removal | C# delete cell prefix Aspose
// Common Searches: how to remove leading apostrophe Aspose.Cells C# | Style.QuotePrefix false example | apply StyleFlag to cell Aspose | remove single quote from Excel cell programmatically | Aspose.Cells delete apostrophe prefix
// Developer Intent: Eliminate the leading apostrophe of a cell by setting Style.QuotePrefix to false and applying it with a StyleFlag.
// Use Cases: Convert text values entered with a leading apostrophe into true numbers for calculations. | Clean imported CSV or user‑entered data where the apostrophe prevents proper sorting or filtering. | Prepare worksheets for export to systems that reject the leading single‑quote character.
// AI Prompts: Generate C# code that opens an existing workbook, removes apostrophe prefixes from an entire column using Style.QuotePrefix = false with a StyleFlag, and saves the file. | Explain the relationship between Style, StyleFlag, and QuotePrefix in Aspose.Cells and why setting QuotePrefix to false removes the leading apostrophe. | Create a reusable method in C# that accepts a worksheet and a range, then removes leading apostrophes from all cells in that range using Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to delete the leading single‑quote in a cell (e.g., B2) by creating a Style with QuotePrefix set to false, enabling the property with a StyleFlag, applying it to the cell, and saving the workbook as RemovedApostrophe.xlsx.
class RemoveApostropheDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put a value that starts with a single quote (apostrophe)
        Cell cell = worksheet.Cells["B2"];
        cell.PutValue("'12345");

        // Create a style and set QuotePrefix to false (remove the leading apostrophe)
        Style style = workbook.CreateStyle();
        style.QuotePrefix = false;

        // Create a StyleFlag to indicate that the QuotePrefix property should be applied
        StyleFlag flag = new StyleFlag();
        flag.QuotePrefix = true;

        // Apply the style to the cell using the flag
        cell.SetStyle(style, flag);

        // Save the workbook
        workbook.Save("RemovedApostrophe.xlsx");
    }
}
