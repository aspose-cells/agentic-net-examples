// Title: How to retrieve the RGB values of the Hyperlink theme color from an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells get hyperlink theme color RGB from workbook | retrieve Excel hyperlink theme color components with Aspose.Cells .NET | Aspose.Cells Workbook.GetThemeColor Hyperlink RGB example | read theme color for hyperlinks in a loaded Excel file using Aspose.Cells | extract hyperlink theme color values from an .xlsx file in C#
// Tags: Workbook.GetThemeColor Hyperlink | Aspose.Cells theme color extraction | Excel hyperlink theme RGB Aspose | C# retrieve theme color from workbook | Aspose.Cells GetThemeColor usage

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The code loads an Excel workbook (creating a simple one if it doesn't exist), optionally handling password‑protected files, then uses Workbook.GetThemeColor with ThemeColorType.Hyperlink to obtain the hyperlink theme color as a System.Drawing.Color and outputs its red, green, and blue components.
class Program
{
    static void Main(string[] args)
    {
        const string filePath = "input.xlsx";

        try
        {
            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(filePath))
            {
                var wb = new Workbook();
                wb.Worksheets[0].Name = "Sheet1";
                wb.Save(filePath);
            }

            Workbook workbook = null;

            // Try to load the workbook without a password.
            try
            {
                workbook = new Workbook(filePath);
            }
            catch (Exception)
            {
                // If loading fails (e.g., password‑protected), retry with an empty password.
                var loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = "" };
                workbook = new Workbook(filePath, loadOptions);
            }

            // Retrieve the Hyperlink theme color.
            Color hyperlinkColor = workbook.GetThemeColor(ThemeColorType.Hyperlink);

            // Extract RGB components.
            int red = hyperlinkColor.R;
            int green = hyperlinkColor.G;
            int blue = hyperlinkColor.B;

            Console.WriteLine($"Hyperlink Theme Color RGB: ({red}, {green}, {blue})");
        }
        catch (Exception e)
        {
            Console.WriteLine($"Error: {e.Message}");
        }
    }
}
