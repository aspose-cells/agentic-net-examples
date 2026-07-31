// Title: Merge Header Row Cells and Set 14‑Point Font with Aspose.Cells for .NET (C#)
// Description: Loads an existing workbook, merges cells A1‑D1 into a single header cell, changes the font size of that merged cell to 14 points, and saves the result as a new Excel file using Aspose.Cells for C#.
// Keywords: Aspose.Cells C# merge cells | Aspose.Cells set font size | merge header row Aspose.Cells | format header cell .NET | export Excel Aspose.Cells | C# Excel cell merging | Aspose.Cells style header
// Common Searches: Aspose.Cells merge first row cells C# | How to set font size for merged cell Aspose.Cells | C# merge header row and change font Aspose.Cells | Save workbook after merging cells Aspose.Cells .NET | Aspose.Cells format header row Excel
// Developer Intent: Combine cells A1‑D1 into one header, apply a 14‑point font, and save the workbook.
// Use Cases: Create a centered title across columns A‑D for financial reports with a larger font before exporting. | Design printable invoice headers by merging cells and enlarging the font, then saving the file for distribution. | Standardize template worksheets by programmatically merging header cells and applying a consistent font size across all reports. | Prepare dashboard sheets where the main heading spans multiple columns and needs a specific typographic style.
// AI Prompts: Generate C# code using Aspose.Cells to merge cells A1:E1, set the font to 14 bold, center the text, and export the workbook as PDF. | Show an example that merges the header row, applies a custom style (font size, alignment, background color), and writes the workbook to a memory stream. | Explain how to undo a cell merge or adjust column widths after merging with Aspose.Cells in C#.

using System;
using Aspose.Cells;

// Loads an existing workbook, merges cells A1‑D1 into a single header cell, changes the font size of that merged cell to 14 points, and saves the result as a new Excel file using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        // Load the existing Excel file
        Workbook workbook = new Workbook("input.xlsx"); // lifecycle: load
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge the header row cells (e.g., A1:D1)
        // Parameters: firstRow (0‑based), firstColumn (0‑based), totalRows (1‑based), totalColumns (1‑based)
        cells.Merge(0, 0, 1, 4); // merges cells A1, B1, C1, D1 into a single cell

        // Set the font size of the merged header cell to 14 points
        Style headerStyle = cells[0, 0].GetStyle();
        headerStyle.Font.Size = 14;
        cells[0, 0].SetStyle(headerStyle);

        // Save the modified workbook
        workbook.Save("output.xlsx"); // lifecycle: save
    }
}
