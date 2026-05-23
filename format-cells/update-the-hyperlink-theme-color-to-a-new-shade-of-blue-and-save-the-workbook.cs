using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

class UpdateHyperlinkThemeColor
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define a new shade of blue for the hyperlink theme color (DodgerBlue)
            Color newBlue = Color.FromArgb(30, 144, 255);
            workbook.SetThemeColor(ThemeColorType.Hyperlink, newBlue);

            // Get the first worksheet and target cell
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Aspose.Cells Hyperlink");

            // Create a style that uses the Hyperlink theme color
            Style style = workbook.CreateStyle();
            style.Font.ThemeColor = new ThemeColor(ThemeColorType.Hyperlink, 0);
            style.Font.Underline = FontUnderlineType.Single;
            cell.SetStyle(style);

            // Add a hyperlink to the cell (A1)
            // Note: Add method signature expects row/column parameters before the address in some versions
            sheet.Hyperlinks.Add(0, 0, 1, 1, "https://www.aspose.com");

            // Save the workbook
            string outputPath = "HyperlinkThemeColorUpdated.xlsx";

            // Ensure the output directory exists (handle case when outputPath has no directory part)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}