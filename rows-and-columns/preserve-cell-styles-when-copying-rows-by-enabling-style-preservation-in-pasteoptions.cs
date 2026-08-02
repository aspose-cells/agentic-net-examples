// Title: Copy a Row with Full Formatting Using PasteOptions in Aspose.Cells for .NET (C#)
// Description: Shows how to duplicate a styled header row to another position while preserving fonts, colors, borders, and other formatting by using Cells.CopyRows with CopyOptions (ExtendToAdjacentRange) and PasteOptions set to PasteType.All in Aspose.Cells for .NET.
// Keywords: Aspose.Cells CopyRows | preserve row style | PasteOptions PasteType.All | C# copy row formatting | Aspose.Cells .NET styling | CopyRows with formatting | Excel row style duplication | Aspose.Cells PasteOptions example | CopyRows PreserveFormatting | CopyOptions ExtendToAdjacentRange
// Common Searches: Aspose.Cells copy row keep formatting | How to preserve cell styles when copying rows in Aspose.Cells C# | PasteOptions PasteType.All example | CopyRows with style preservation .NET | Aspose.Cells CopyOptions ExtendToAdjacentRange usage
// Developer Intent: Duplicate a row while retaining all cell styles and formatting.
// Use Cases: Copy a formatted header row into a newly inserted row so column titles keep bold text, white font, and dark‑blue background. | Insert a blank row in a financial report and replicate a template row with its borders and shading intact. | Generate multiple data‑entry rows that share identical styling by copying a single styled row to several destinations.
// AI Prompts: Write C# code using Aspose.Cells to copy a row and preserve its style with PasteOptions.PasteType.All. | Explain how CopyOptions.ExtendToAdjacentRange and PasteOptions affect formatting when copying rows in Aspose.Cells. | Provide a step‑by‑step guide for copying a styled row while keeping fonts, colors, and borders unchanged in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to duplicate a styled header row to another position while preserving fonts, colors, borders, and other formatting by using Cells.CopyRows with CopyOptions (ExtendToAdjacentRange) and PasteOptions set to PasteType.All in Aspose.Cells for .NET.
    class PreserveRowStylesExample
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sourceSheet = workbook.Worksheets[0];
                Cells sourceCells = sourceSheet.Cells;

                // Populate header data in the first row
                sourceCells["A1"].PutValue("Header 1");
                sourceCells["B1"].PutValue("Header 2");
                sourceCells["C1"].PutValue("Header 3");

                // Create a style for the header row
                Style rowStyle = workbook.CreateStyle();
                rowStyle.Font.IsBold = true;
                rowStyle.Font.Color = Color.White;
                rowStyle.ForegroundColor = Color.DarkBlue;
                rowStyle.Pattern = BackgroundType.Solid;

                // Apply the style to the entire first row
                Row headerRow = sourceSheet.Cells.Rows[0];
                headerRow.ApplyStyle(rowStyle, new StyleFlag { All = true });

                // Insert a blank row at index 5 where the source row will be copied
                sourceSheet.Cells.InsertRows(5, 1);

                // Set copy options (extend to adjacent range)
                CopyOptions copyOptions = new CopyOptions
                {
                    ExtendToAdjacentRange = true
                };

                // Set paste options to preserve all data including formats
                PasteOptions pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.All
                };

                // Copy the first row (index 0) to the newly inserted row (index 5)
                sourceSheet.Cells.CopyRows(
                    sourceSheet.Cells,   // source cells
                    0,                   // source row index
                    5,                   // destination row index
                    1,                   // number of rows to copy
                    copyOptions,
                    pasteOptions);

                // Save the workbook
                workbook.Save("PreserveRowStylesOutput.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
