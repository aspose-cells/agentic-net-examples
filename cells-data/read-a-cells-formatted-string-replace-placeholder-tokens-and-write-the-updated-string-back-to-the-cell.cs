// Title: Aspose.Cells for .NET – Read formatted cell text, replace placeholders, and write back
// Description: Demonstrates how to load or create an Excel workbook with Aspose.Cells, read a cell's formatted string, substitute placeholder tokens (e.g., {Name}, {OrderId}) with actual values, write the updated text to the same cell, and save the file while preserving cell formatting.
// Keywords: Aspose.Cells placeholder replacement | read formatted string cell Aspose | write updated text Excel .NET | C# Excel token substitution | Aspose.Cells string replace | Excel cell text update Aspose | placeholder substitution Aspose.Cells
// Common Searches: replace tokens in Excel cell using Aspose.Cells C# | read cell string value and update it Aspose.Cells | Aspose.Cells replace {Name} placeholder | how to write modified text back to same cell Aspose | preserve cell formatting while replacing text Aspose.Cells
// Developer Intent: Swap placeholder tokens in a cell's text and save the modified value with Aspose.Cells for .NET.
// Use Cases: Generate personalized letters by inserting customer names and order IDs into template strings stored in worksheet cells. | Update dynamic status messages in reports by replacing placeholders with runtime data before exporting the workbook. | Batch‑process an existing spreadsheet to replace placeholders across multiple cells and overwrite the original content.
// AI Prompts: Show C# code with Aspose.Cells that reads a cell's string, replaces placeholders like {Date} and {Amount}, and writes the result back to the same cell. | Provide an example that iterates over a range of cells in Aspose.Cells and performs placeholder substitution for each cell's text. | Explain how to keep number/date formatting intact while replacing placeholder tokens in an Excel cell using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsPlaceholderReplacement
{
    // Demonstrates how to load or create an Excel workbook with Aspose.Cells, read a cell's formatted string, substitute placeholder tokens (e.g., {Name}, {OrderId}) with actual values, write the updated text to the same cell, and save the file while preserving cell formatting.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // new workbook
            // If you need to load an existing file, use:
            // Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Example: put a formatted string with a placeholder into cell A1
            Cell targetCell = worksheet.Cells["A1"];
            targetCell.PutValue("Dear {Name}, your order {OrderId} is confirmed.");

            // Read the cell's formatted string (includes any number/date formatting)
            string formattedString = targetCell.StringValue;

            // Replace placeholder tokens with actual values
            string replacedString = formattedString
                .Replace("{Name}", "John Doe")
                .Replace("{OrderId}", "12345");

            // Write the updated string back to the same cell
            targetCell.PutValue(replacedString);

            // Save the workbook to a file
            workbook.Save("Output.xlsx", SaveFormat.Xlsx);
        }
    }
}
