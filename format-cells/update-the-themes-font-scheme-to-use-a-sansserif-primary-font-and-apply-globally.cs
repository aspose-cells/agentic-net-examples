// Title: Apply a Global Sans‑Serif Font (Arial) to an Aspose.Cells Workbook in C#
// Description: Shows how to change the workbook's default style to Arial (12 pt) using Aspose.Cells for .NET, apply the style to a cell, and save the file, providing a quick way to enforce a consistent font across all worksheets.
// Keywords: Aspose.Cells default font C# | global font scheme Aspose.Cells | set workbook default style .NET | Arial font Excel Aspose | apply sans-serif font workbook | theme font update Aspose.Cells | C# Excel default style | change workbook font programmatically
// Common Searches: change default font in Aspose.Cells workbook C# | set global font for all cells Aspose.Cells .NET | apply Arial as default style in Excel using Aspose | update theme font scheme in Aspose.Cells C# | C# code to set workbook default font
// Developer Intent: Configure the workbook so every cell automatically uses a specified sans‑serif font.
// Use Cases: Create reports that must follow a corporate Arial style without manual formatting. | Build a template workbook where any added data inherits the chosen font. | Generate Excel files for publishing where a consistent sans‑serif appearance is required.
// AI Prompts: Generate C# code to change an Aspose.Cells workbook’s default font to Helvetica and apply it to existing cells. | Explain how to modify the ThemeFontScheme in Aspose.Cells so both primary and secondary fonts become a custom sans‑serif typeface. | Provide a step‑by‑step guide for setting a global font in an Excel file using Aspose.Cells for .NET without styling each cell individually.

using System;
using Aspose.Cells;

namespace ThemeFontSchemeExample
{
    // Shows how to change the workbook's default style to Arial (12 pt) using Aspose.Cells for .NET, apply the style to a cell, and save the file, providing a quick way to enforce a consistent font across all worksheets.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Set the global default font for the workbook
                workbook.DefaultStyle.Font.Name = "Arial";
                workbook.DefaultStyle.Font.Size = 12;

                // Create a style that uses the desired font (optional, same as default)
                Style globalStyle = workbook.CreateStyle();
                globalStyle.Font.Name = "Arial";
                globalStyle.Font.Size = 12;

                // Apply the style to a sample cell
                Worksheet sheet = workbook.Worksheets[0];
                Cell sampleCell = sheet.Cells["A1"];
                sampleCell.PutValue("Text using the global sans‑serif theme font");
                sampleCell.SetStyle(globalStyle);

                // Save the workbook
                string outputPath = "ThemeFontSchemeUpdated.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
