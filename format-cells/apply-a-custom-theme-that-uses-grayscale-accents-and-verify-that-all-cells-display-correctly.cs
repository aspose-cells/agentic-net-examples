// Title: Apply a custom grayscale font style to a cell range and verify the result using Aspose.Cells for .NET
// AI Prompts: Create a dark gray Font style, assign it to a Style object, and apply it to the range A1:A3 in a new workbook with Aspose.Cells. | Save the workbook as GrayscaleTheme.xlsx, then open the file again and retrieve the style of cell A1. | Compare the retrieved font color with the original gray value and output whether they match. | Handle missing‑file scenarios and display a clear verification message.
// Common Searches: how to set a custom gray font color for a range of cells using Aspose.Cells C# | verify font color of a cell after saving and reloading an Excel workbook with Aspose.Cells | apply grayscale theme to specific cells in an XLSX file using Aspose.Cells for .NET | check if a cell uses a particular RGB color after workbook reload in C# Aspose.Cells
// Tags: apply grayscale font style Aspose.Cells | range.ApplyStyle set font color C# | verify cell font color after workbook reload Aspose.Cells | save workbook as XLSX with custom theme Aspose.Cells | create custom color style Aspose.Cells .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example creates a new workbook, defines a dark gray color, builds a font style with that color, applies the style to cells A1‑A3, saves the file as GrayscaleTheme.xlsx, reloads the workbook, reads the font color of cell A1, and prints whether the loaded color matches the defined gray.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Define grayscale color (used later for styling)
            Color accentGray = Color.FromArgb(0xFF, 0x33, 0x33, 0x33); // Dark gray

            // Populate some cells
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Header");
            ws.Cells["A2"].PutValue(123);
            ws.Cells["A3"].PutValue(456);

            // Create a style that uses the defined gray for the font
            Style accentStyle = wb.CreateStyle();
            accentStyle.Font.Color = accentGray;
            accentStyle.Font.Size = 12;

            // Apply the style to the range A1:A3
            AsposeRange range = ws.Cells.CreateRange("A1:A3");
            range.ApplyStyle(accentStyle, new StyleFlag { Font = true });

            // Save the workbook
            string filePath = "GrayscaleTheme.xlsx";
            wb.Save(filePath);

            // ---- Verification ----
            // Reload the workbook only if the file exists
            if (File.Exists(filePath))
            {
                try
                {
                    Workbook wbVerify = new Workbook(filePath);
                    Worksheet wsVerify = wbVerify.Worksheets[0];
                    Style verifyStyle = wsVerify.Cells["A1"].GetStyle();

                    bool usesGray = verifyStyle.Font.Color.ToArgb() == accentGray.ToArgb();
                    Console.WriteLine("Verification - Font uses defined gray: " + usesGray);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Verification error: " + ex.Message);
                }
            }
            else
            {
                Console.WriteLine("Verification failed: file not found.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
