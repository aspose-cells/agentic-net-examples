// Title: Replace commas with semicolons in a cell's formatted text and save the workbook using Aspose.Cells for .NET
// AI Prompts: Read the displayed string from a target cell, substitute every ',' with ';', and write the updated text back to the same cell with Aspose.Cells in C#. | Load an Excel file, obtain a cell's formatted value, perform a comma‑to‑semicolon replacement, update the cell, and persist the changes using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# replace commas with semicolons in a specific cell | how to get and modify formatted cell text with Aspose.Cells .NET | write updated string back to the same Excel cell using Aspose.Cells | replace characters in Excel cell value and save workbook Aspose.Cells | comma to semicolon substitution in cell A1 using Aspose.Cells C#
// Tags: comma to semicolon replacement cell Aspose.Cells | formatted cell text update .NET | modify cell value and save workbook Aspose.Cells | character substitution Excel cell C# | cell string manipulation Aspose.Cells API

using System;
using Aspose.Cells;

// The example loads an Excel workbook, reads the formatted text of a specified cell, replaces all commas with semicolons, writes the modified string back to the same cell, and saves the workbook.
class ReplaceCommasInCell
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or specify the desired one)
        Worksheet worksheet = workbook.Worksheets[0];

        // Choose the cell to process (e.g., A1)
        Cell cell = worksheet.Cells["A1"];

        // Retrieve the cell's formatted text
        string formattedText = cell.StringValue;

        // Replace commas with semicolons
        string modifiedText = formattedText.Replace(',', ';');

        // Write the modified string back to the same cell
        cell.PutValue(modifiedText);

        // Save the workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
