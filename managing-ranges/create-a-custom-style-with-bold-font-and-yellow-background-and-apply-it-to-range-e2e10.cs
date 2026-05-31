using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsCustomStyleDemo
{
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

                // Create a custom style with bold font and yellow background
                Style customStyle = workbook.CreateStyle();
                customStyle.Font.IsBold = true;                     // Bold font
                customStyle.Pattern = BackgroundType.Solid;         // Enable solid fill
                customStyle.ForegroundColor = Color.Yellow;         // Yellow background

                // Define the range E2:E10 (zero‑based indices: row 1, column 4, 9 rows, 1 column)
                Aspose.Cells.Range targetRange = cells.CreateRange(1, 4, 9, 1);

                // Apply the custom style to the entire range
                targetRange.SetStyle(customStyle);

                // Save the workbook
                workbook.Save("CustomStyle_E2_E10.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}