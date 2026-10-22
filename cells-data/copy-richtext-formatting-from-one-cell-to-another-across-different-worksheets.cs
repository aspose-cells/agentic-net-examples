// Title: Copy rich‑text formatting and style from a cell in one worksheet to another cell in a different worksheet using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells for C# to duplicate the rich‑text content and full formatting of cell A1 on Sheet1 into cell B2 on Sheet2, then save the workbook. | Write C# code that copies the style (font, color, underline, etc.) from a source cell to a destination cell on another worksheet with Aspose.Cells, preserving the original text. | Modify the example to also transfer merged‑cell ranges and cell comments while copying formatting between worksheets using Aspose.Cells.
// Common Searches: Aspose.Cells copy cell formatting from one sheet to another in C# | How to preserve rich text style when moving a cell between worksheets using Aspose.Cells .NET | C# example for transferring cell style and value across different worksheets with Aspose.Cells | Copy merged cells and comments along with formatting between sheets using Aspose.Cells
// Tags: Aspose.Cells transfer cell style between worksheets | C# preserve rich text formatting Aspose.Cells | copy merged cell ranges Aspose.Cells | transfer cell comments Aspose.Cells C# | apply source cell style to destination cell Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads a workbook, copies the string value and full rich‑text style from cell A1 on Sheet1 to cell B2 on Sheet2, applies the source cell’s style to the destination, and saves the result as output.xlsx, with error handling for missing files or worksheets.
class RichTextCopyExample
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(sourcePath);

            // Access the source and destination worksheets (adjust names if necessary)
            Worksheet sourceSheet = workbook.Worksheets["Sheet1"];
            Worksheet destinationSheet = workbook.Worksheets["Sheet2"];

            if (sourceSheet == null || destinationSheet == null)
            {
                Console.WriteLine("One or both specified worksheets were not found.");
                return;
            }

            // Define the source and destination cells
            Cell sourceCell = sourceSheet.Cells["A1"];          // cell with rich‑text to copy
            Cell destinationCell = destinationSheet.Cells["B2"]; // target cell

            // Copy plain text value from source to destination
            destinationCell.PutValue(sourceCell.StringValue);

            // Copy the entire cell style (includes font, color, etc.)
            destinationCell.SetStyle(sourceCell.GetStyle());

            // Save the workbook with the changes
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
