// Title: How to copy a cell’s style and keep its borders with Aspose.Cells StyleFlag in C#
// Common Searches: Aspose.Cells copy cell formatting with borders C# example | Preserve Excel cell borders when applying style from another cell using Aspose.Cells | SetStyle with StyleFlag to transfer borders in .NET | Copy style including borders from A1 to B1 Aspose.Cells C# | How to retain cell border styles while copying style in Aspose.Cells for .NET
// Tags: Aspose.Cells SetStyle StyleFlag usage | style transfer with borders Aspose.Cells | retain cell border formatting Aspose.Cells | StyleFlag All property for formatting transfer | apply source cell formatting to target cell Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The example loads an Excel workbook (creating a simple one with a bordered A1 cell if needed), extracts the style from cell A1, creates a StyleFlag with All set to true, and applies that style to cell B1 using SetStyle. This copies the complete formatting—including all border settings—to the target cell, then saves the result as output.xlsx.
class PreserveBorders
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; if not, create a simple workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("Source");

                // Apply a sample border to A1 so we can see it being preserved.
                Style srcStyle = ws.Cells["A1"].GetStyle();
                srcStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
                srcStyle.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
                srcStyle.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                srcStyle.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
                ws.Cells["A1"].SetStyle(srcStyle);

                // Save the template workbook.
                workbook.Save(inputPath);
            }

            // Get the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Define source and destination cells.
            Cell sourceCell = worksheet.Cells["A1"];
            Cell targetCell = worksheet.Cells["B1"];

            // Retrieve the style from the source cell.
            Style sourceStyle = sourceCell.GetStyle();

            // Create a StyleFlag that includes all style elements (including borders).
            StyleFlag styleFlag = new StyleFlag { All = true };

            // Apply the source style to the target cell while preserving borders.
            targetCell.SetStyle(sourceStyle, styleFlag);

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
