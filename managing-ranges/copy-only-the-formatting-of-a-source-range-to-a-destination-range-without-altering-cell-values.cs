using System;
using System.Drawing;
using Aspose.Cells;

class CopyFormattingDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // ---------- Source worksheet ----------
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "Source";

            // Fill source range with data
            Cells sourceCells = sourceSheet.Cells;
            Aspose.Cells.Range sourceRange = sourceCells.CreateRange("A1:C3");
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    sourceCells[row, col].PutValue($"Src{row + 1}{col + 1}");
                }
            }

            // Apply a style to the source range
            Style srcStyle = workbook.CreateStyle();
            srcStyle.Font.Name = "Arial";
            srcStyle.Font.Size = 12;
            srcStyle.Font.IsBold = true;
            srcStyle.ForegroundColor = Color.LightBlue;
            srcStyle.Pattern = BackgroundType.Solid;
            sourceRange.SetStyle(srcStyle);

            // ---------- Destination worksheet ----------
            Worksheet destSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            destSheet.Name = "Destination";

            // Fill destination range with different values
            Cells destCells = destSheet.Cells;
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    destCells[row, col].PutValue($"Dst{row + 1}{col + 1}");
                }
            }

            // Create destination range
            Aspose.Cells.Range destRange = destCells.CreateRange("A1:C3");

            // Copy only the formatting from source range to destination range
            destRange.CopyStyle(sourceRange);

            // Save the workbook
            workbook.Save("CopyFormattingDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}