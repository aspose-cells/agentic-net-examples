using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

class ApplyHyperlinkThemeColor
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.ods";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the ODS workbook with the option to apply Excel default hyperlink style
            OdsLoadOptions loadOptions = new OdsLoadOptions
            {
                ApplyExcelDefaultStyleToHyperlink = true
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Iterate through all worksheets and their hyperlinks
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                for (int i = 0; i < sheet.Hyperlinks.Count; i++)
                {
                    Hyperlink link = sheet.Hyperlinks[i];

                    // Convert the hyperlink's CellArea to a Range object
                    CellArea area = link.Area;
                    int totalRows = area.EndRow - area.StartRow + 1;
                    int totalColumns = area.EndColumn - area.StartColumn + 1;
                    Aspose.Cells.Range range = sheet.Cells.CreateRange(area.StartRow, area.StartColumn, totalRows, totalColumns);

                    // Create a style that uses the theme's hyperlink color (commonly blue) and underline
                    Style hyperlinkStyle = workbook.CreateStyle();
                    hyperlinkStyle.Font.Color = Color.Blue;               // Theme hyperlink color
                    hyperlinkStyle.Font.Underline = FontUnderlineType.Single;

                    // Apply the style to the hyperlink range
                    range.SetStyle(hyperlinkStyle);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}