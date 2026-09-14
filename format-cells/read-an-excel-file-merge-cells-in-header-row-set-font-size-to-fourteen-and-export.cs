// Title: How to merge header cells A1:D1, apply a 14‑point bold centered style, and save the workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing .xlsx, merge cells A1 through D1, set the font to 14‑point bold and center it, then save the file with Aspose.Cells in C#. | Write a C# program that opens input.xlsx, merges the first row across four columns, applies a centered 14‑point bold style to the merged range, and writes the result to output.xlsx using Aspose.Cells.
// Common Searches: Aspose.Cells C# merge first row cells and set font size to 14 | C# code to merge A1:D1 and apply bold centered style in Excel workbook | How to style merged header cells with Aspose.Cells for .NET | Save modified Excel file after merging cells using Aspose.Cells C#
// Tags: merge cells A1:D1 Aspose.Cells C# | apply 14 point bold centered style Aspose.Cells | format merged header range Aspose.Cells | load and edit existing .xlsx Aspose.Cells | save workbook as new file Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.IO;

// // Loads input.xlsx, merges cells A1‑D1 on the first worksheet, applies a 14‑point bold centered style to the merged header, and saves the result as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (you can change the index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range of header cells to merge (e.g., A1 to D1)
            CellArea headerArea = new CellArea
            {
                StartRow = 0,      // Row 1 (0‑based)
                EndRow = 0,        // Row 1
                StartColumn = 0,   // Column A (0‑based)
                EndColumn = 3      // Column D
            };

            // Merge the defined header cells
            sheet.Cells.Merge(
                headerArea.StartRow,
                headerArea.StartColumn,
                headerArea.EndRow - headerArea.StartRow + 1,
                headerArea.EndColumn - headerArea.StartColumn + 1);

            // Create a style for the merged header
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.Size = 14;
            headerStyle.Font.IsBold = true;
            headerStyle.HorizontalAlignment = TextAlignmentType.Center;
            headerStyle.VerticalAlignment = TextAlignmentType.Center;

            // Apply the style to the merged cells
            Aspose.Cells.Range headerRange = sheet.Cells.CreateRange(
                headerArea.StartRow,
                headerArea.StartColumn,
                headerArea.EndRow - headerArea.StartRow + 1,
                headerArea.EndColumn - headerArea.StartColumn + 1);

            headerRange.ApplyStyle(
                headerStyle,
                new StyleFlag { Font = true, HorizontalAlignment = true, VerticalAlignment = true });

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
