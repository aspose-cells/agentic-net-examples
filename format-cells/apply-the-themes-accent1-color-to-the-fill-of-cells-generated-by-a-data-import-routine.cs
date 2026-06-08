using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Simulate a data import routine: fill column A with labels and column B with values
            for (int row = 0; row < 10; row++)
            {
                worksheet.Cells[row, 0].PutValue($"Item {row + 1}");
                worksheet.Cells[row, 1].PutValue(row * 10);
            }

            // Create a style that uses the theme's Accent1 color for the cell fill
            Style accentStyle = workbook.CreateStyle();
            accentStyle.Pattern = BackgroundType.Solid; // Enable solid fill
            accentStyle.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent1, 0); // Accent1 with no tint

            // Apply the style to the imported data range
            AsposeRange dataRange = worksheet.Cells.CreateRange("A1:B10");
            StyleFlag flag = new StyleFlag();
            flag.CellShading = true; // Ensure fill is applied
            dataRange.ApplyStyle(accentStyle, flag);

            // Save the workbook
            string outputPath = "DataWithAccent1Fill.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display the error
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}