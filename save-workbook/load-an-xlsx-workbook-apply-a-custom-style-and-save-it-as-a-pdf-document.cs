using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Create a custom style
            Style customStyle = workbook.CreateStyle();
            customStyle.ForegroundColor = Color.LightYellow;          // Background color
            customStyle.Pattern = BackgroundType.Solid;               // Apply solid fill
            customStyle.Font.Name = "Arial";                          // Font name
            customStyle.Font.Size = 12;                               // Font size
            customStyle.Font.IsBold = true;                           // Bold text

            // Define the range to which the style will be applied (A1:C5)
            Worksheet worksheet = workbook.Worksheets[0];
            CellArea area = CellArea.CreateCellArea("A1", "C5");

            // Create a range object for the defined area
            int startRow = area.StartRow;
            int startColumn = area.StartColumn;
            int totalRows = area.EndRow - area.StartRow + 1;
            int totalColumns = area.EndColumn - area.StartColumn + 1;
            Aspose.Cells.Range range = worksheet.Cells.CreateRange(startRow, startColumn, totalRows, totalColumns);

            // Apply the style to the range
            StyleFlag flag = new StyleFlag { All = true };
            range.ApplyStyle(customStyle, flag);

            // Save the modified workbook as PDF
            string pdfPath = "output.pdf";
            try
            {
                workbook.Save(pdfPath, SaveFormat.Pdf);
                Console.WriteLine($"Workbook saved as PDF: {pdfPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving PDF: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}