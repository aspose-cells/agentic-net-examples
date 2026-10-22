// Title: How to apply a custom cell style to every cell in a named range using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# apply custom style to a named range in an existing workbook | set font fill and borders for all cells of a named range with Aspose.Cells | retrieve named range address and format its cells programmatically in .NET | using StyleFlag.All to format a range of cells in Aspose.Cells C# | apply style to a named range without hard‑coding the sheet name Aspose.Cells
// Tags: named range style application Aspose.Cells | StyleFlag All usage in .NET | create custom Excel cell style C# | retrieve named range address Aspose.Cells | range.ApplyStyle method example | format cells with borders background Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads or creates an Excel workbook, defines a custom style (font, size, color, fill, and borders), retrieves a named range, resolves its worksheet and address, creates a Range object for that address, and applies the style to every cell in the range using a StyleFlag with All=true, then saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file does not exist.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets.Add("Sheet1");
            }

            // Create a custom style.
            Style style = workbook.CreateStyle();
            style.Font.Name = "Arial";
            style.Font.Size = 12;
            style.Font.Color = Color.Blue;
            style.ForegroundColor = Color.LightYellow;
            style.Pattern = BackgroundType.Solid;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].Color = Color.Black;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].Color = Color.Black;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].Color = Color.Black;
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.TopBorder].Color = Color.Black;

            // Apply all style attributes.
            StyleFlag flag = new StyleFlag { All = true };

            // Retrieve the named range (replace "MyRange" with your actual range name).
            Name namedRange = workbook.Worksheets.Names["MyRange"];
            if (namedRange != null)
            {
                // The RefersTo property contains the address, possibly with sheet name (e.g., "Sheet1!A1:C10").
                string address = namedRange.RefersTo;
                // Separate sheet name and cell range.
                string sheetName;
                string rangeAddress;
                if (address.Contains("!"))
                {
                    var parts = address.Split('!');
                    sheetName = parts[0].Trim('\''); // remove possible quotes
                    rangeAddress = parts[1];
                }
                else
                {
                    // If no sheet name is specified, use the first worksheet.
                    sheetName = workbook.Worksheets[0].Name;
                    rangeAddress = address;
                }

                // Get the worksheet that contains the range.
                Worksheet ws = workbook.Worksheets[sheetName];
                if (ws != null)
                {
                    // Create a Range object for the address.
                    Aspose.Cells.Range range = ws.Cells.CreateRange(rangeAddress);
                    // Apply the custom style to every cell in the range.
                    range.ApplyStyle(style, flag);
                }
            }

            // Save the workbook with the applied formatting.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
