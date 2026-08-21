// Title: Set a Custom Purple Hyperlink Theme and Apply It to All Hyperlink Cells with Aspose.Cells for .NET
// Description: Creates a workbook, defines a purple RGB (128,0,128) as the Hyperlink theme color, adds sample links, forces each hyperlink cell to use the same purple font, and saves the file as HyperlinkThemePurple.xlsx using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | hyperlink theme color | custom purple | SetThemeColor | Excel hyperlink font color | apply color to all hyperlinks | Workbook.Save | SaveFormat.Xlsx
// Common Searches: Aspose.Cells change hyperlink color to purple | C# set hyperlink theme color in Excel | apply custom font color to all hyperlinks Aspose | override hyperlink theme when not supported | update hyperlink cell style programmatically
// Developer Intent: Define a purple Hyperlink theme and ensure every hyperlink cell displays that exact color.
// Use Cases: Brand‑consistent Excel reports that require purple hyperlinks. | Batch updating existing workbooks where theme references are unavailable. | Generating marketing or presentation files with a uniform hyperlink appearance.
// AI Prompts: Generate C# code with Aspose.Cells to set the Hyperlink theme to an RGB value and force all existing hyperlink cells to use the same font color. | Show how to iterate through worksheets and hyperlink ranges to apply a custom color when the theme color cannot be referenced. | Explain why direct font color assignment may be needed after calling Workbook.SetThemeColor for hyperlinks.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHyperlinkThemeDemo
{
    // Creates a workbook, defines a purple RGB (128,0,128) as the Hyperlink theme color, adds sample links, forces each hyperlink cell to use the same purple font, and saves the file as HyperlinkThemePurple.xlsx using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Define a custom purple color for hyperlinks
                Color customPurple = Color.FromArgb(128, 0, 128); // RGB (128,0,128)

                // Apply the custom purple to the Hyperlink theme color
                workbook.SetThemeColor(ThemeColorType.Hyperlink, customPurple);

                // Add sample hyperlinks
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");
                sheet.Hyperlinks.Add("B2", 1, 1, "https://www.example.com");
                sheet.Cells["A1"].PutValue("Aspose");
                sheet.Cells["B2"].PutValue("Example");

                // Ensure each hyperlink cell explicitly uses the custom purple color
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    foreach (Hyperlink hl in ws.Hyperlinks)
                    {
                        CellArea area = hl.Area;

                        for (int row = area.StartRow; row <= area.EndRow; row++)
                        {
                            for (int col = area.StartColumn; col <= area.EndColumn; col++)
                            {
                                Style style = ws.Cells[row, col].GetStyle();
                                // Directly set the font color to the custom purple (theme color reference not available in this version)
                                style.Font.Color = customPurple;
                                ws.Cells[row, col].SetStyle(style);
                            }
                        }
                    }
                }

                // Save the workbook with the updated hyperlink theme color
                workbook.Save("HyperlinkThemePurple.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
