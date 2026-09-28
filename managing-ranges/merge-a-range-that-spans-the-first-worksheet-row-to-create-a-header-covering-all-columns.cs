// Title: Create a centered bold header by merging the first worksheet row across a variable number of columns using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that merges cells from A1 to the column defined by a variable, writes a header string, centers it, makes the font bold, and saves the workbook. | Show how to programmatically merge the entire first row of a worksheet based on a column count parameter, apply a custom style, and export to an .xlsx file using Aspose.Cells. | Provide a reusable method in C# that accepts a worksheet, column count, and header text, merges the top row, sets a centered bold style, and returns the modified workbook.
// Common Searches: aspnet merge first row cells into a single header with Aspose.Cells | c# Aspose.Cells dynamic column count for merged header row | how to apply centered bold style to a merged cell in Aspose.Cells .NET | save workbook after merging header row using Aspose.Cells C# example
// Tags: merge first row cells Aspose.Cells C# | dynamic header merge Aspose.Cells | centered bold style merged cell Aspose.Cells | save merged header workbook Aspose.Cells | worksheet header spanning multiple columns Aspose.Cells

using Aspose.Cells;
using System;

// // Merges the first row (A1 to the column defined by totalColumns) into a single cell, inserts "Report Header", centers the text, makes it bold, and saves the workbook as HeaderMerged.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define how many columns the header should cover (e.g., 10 columns: A to J)
        int totalColumns = 10;

        // Merge the cells in the first row across the defined columns (A1 to J1)
        sheet.Cells.Merge(0, 0, 1, totalColumns);

        // Set the header text in the merged cell
        sheet.Cells[0, 0].PutValue("Report Header");

        // Optional: Apply a simple style to the header
        Style headerStyle = workbook.CreateStyle();
        headerStyle.HorizontalAlignment = TextAlignmentType.Center;
        headerStyle.Font.IsBold = true;
        headerStyle.Font.Size = 14;
        sheet.Cells[0, 0].SetStyle(headerStyle);

        // Save the workbook to a file
        workbook.Save("HeaderMerged.xlsx");
    }
}
