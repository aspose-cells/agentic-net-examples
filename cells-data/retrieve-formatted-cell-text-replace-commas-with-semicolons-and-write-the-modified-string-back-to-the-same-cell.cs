// Title: Aspose.Cells for .NET – Replace commas with semicolons in a cell’s formatted text (C#)
// Description: This example creates a workbook, reads the displayed string of cell A1 using StringValue, swaps every comma for a semicolon, writes the updated text back with PutValue, and saves the file as Output.xlsx.
// Keywords: Aspose.Cells replace commas | C# cell StringValue | modify Excel cell text | comma to semicolon conversion | .NET write cell value | Aspose.Cells formatted text
// Common Searches: Aspose.Cells C# replace commas in cell | How to change cell text from comma to semicolon using Aspose.Cells | Read and write formatted cell value Aspose.Cells .NET | Replace characters in Excel cell with Aspose.Cells API
// Developer Intent: Swap commas for semicolons in a specific cell’s displayed text and persist the change.
// Use Cases: Clean CSV‑style lists stored in a single cell before export | Prepare data for semicolon‑delimited imports without altering other cell attributes | Apply in‑place string transformations while preserving cell formatting
// AI Prompts: Write C# code with Aspose.Cells that reads a cell’s StringValue, replaces all commas with semicolons, and saves the workbook. | Show an Aspose.Cells snippet that iterates over a range and converts commas to semicolons in each cell’s formatted text. | Explain the difference between Cell.StringValue and Cell.Value in Aspose.Cells and advise which to use for text manipulation.

using System;
using Aspose.Cells;

// This example creates a workbook, reads the displayed string of cell A1 using StringValue, swaps every comma for a semicolon, writes the updated text back with PutValue, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Example data: a cell containing commas
        worksheet.Cells["A1"].PutValue("Apple, Banana, Cherry");

        // Access the target cell
        Cell cell = worksheet.Cells["A1"];

        // Retrieve the formatted text of the cell
        string formattedText = cell.StringValue;

        // Replace commas with semicolons
        string modifiedText = formattedText.Replace(",", ";");

        // Write the modified string back to the same cell
        cell.PutValue(modifiedText);

        // Save the workbook
        workbook.Save("Output.xlsx");
    }
}
