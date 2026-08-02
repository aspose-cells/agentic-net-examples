// Title: Apply a custom percentage format to column C after offsetting a range by two rows – Aspose.Cells for .NET
// Description: Creates a workbook, defines a range starting at C1, shifts it down two rows, retrieves the EntireColumn of the shifted range (still column C), builds a style with the custom format "0.00%", and applies the number format to the whole column before saving the file.
// Keywords: Aspose.Cells C# custom number format | apply style to entire column | EntireColumn property Aspose.Cells | offset range rows Aspose.Cells | percentage format .NET | column C formatting Aspose.Cells
// Common Searches: Aspose.Cells apply custom number format to a column | How to use EntireColumn after offsetting a range in C# | Set percentage format for column C with Aspose.Cells | Shift range by rows and style column in .NET | Apply style flag number format Aspose.Cells
// Developer Intent: Format column C with a two‑decimal percentage style after moving the source range down two rows.
// Use Cases: Display financial ratios as percentages when data begins at row 3. | Maintain consistent column formatting after inserting header rows. | Quickly reapply a predefined style to a column when the data range is programmatically repositioned.
// AI Prompts: Generate C# code that creates a custom "0.00%" number format and applies it to the EntireColumn of a range offset by two rows using Aspose.Cells. | Explain step‑by‑step how to offset a range, retrieve its EntireColumn, and apply only the number‑format flag in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Creates a workbook, defines a range starting at C1, shifts it down two rows, retrieves the EntireColumn of the shifted range (still column C), builds a style with the custom format "0.00%", and applies the number format to the whole column before saving the file.
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

            // Define the original range (starting at row 0, column 2 i.e., C1)
            int originalStartRow = 0;
            int originalStartColumn = 2; // Column C
            int rowCount = 5;   // number of rows in the original range
            int columnCount = 1; // number of columns in the original range

            // Offset the original range by two rows
            AsposeRange offsetRange = cells.CreateRange(originalStartRow + 2, originalStartColumn, rowCount, columnCount);

            // Get the entire column that contains the offset range (still column C)
            AsposeRange entireColumn = offsetRange.EntireColumn;

            // Create a style with a custom number format (percentage with two decimals)
            Style style = workbook.CreateStyle();
            style.Custom = "0.00%";

            // Configure the style flag to apply only the number format
            StyleFlag styleFlag = new StyleFlag();
            styleFlag.NumberFormat = true;

            // Apply the style to the whole column
            entireColumn.ApplyStyle(style, styleFlag);

            // Save the workbook
            workbook.Save("CustomNumberFormatColumnC.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
