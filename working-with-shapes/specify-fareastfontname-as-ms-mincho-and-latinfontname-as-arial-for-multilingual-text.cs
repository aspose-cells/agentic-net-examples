// Title: Set FarEast and Latin fonts for multilingual shape text with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a rectangle shape, inserts English‑Chinese‑Japanese text, and uses TextOptions.FarEastName = "MS Mincho" and TextOptions.LatinName = "Arial" to render CJK characters and Latin characters with appropriate fonts before saving.
// Keywords: Aspose.Cells | C# | .NET | shape text font | FarEastName | LatinName | multilingual Excel | MS Mincho | Arial | CJK font in Excel
// Common Searches: Aspose.Cells set FarEastName for Japanese text | How to apply different fonts to CJK and Latin characters in a shape | C# Aspose.Cells multilingual shape example | Change font of Chinese characters in Excel shape | Rectangle shape font settings Aspose.Cells .NET
// Developer Intent: Apply distinct fonts to Far‑East (CJK) and Latin scripts within a shape’s text using Aspose.Cells.
// Use Cases: Generate a report where a shape displays English headings in Arial and Japanese subtitles in MS Mincho. | Create marketing dashboards with mixed English and Chinese labels, each rendered in its native font. | Automate Excel templates that require proper font rendering for multilingual annotations inside shapes.
// AI Prompts: Write C# code with Aspose.Cells to set FarEastName to "SimSun" and LatinName to "Calibri" for a shape containing mixed Chinese and English text. | Show how to assign different fonts to Latin and CJK scripts for multiple shapes across a worksheet using Aspose.Cells for .NET. | Explain the impact of TextOptions.FarEastName and TextOptions.LatinName on rendering multilingual text in Excel shapes.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a rectangle shape, inserts English‑Chinese‑Japanese text, and uses TextOptions.FarEastName = "MS Mincho" and TextOptions.LatinName = "Arial" to render CJK characters and Latin characters with appropriate fonts before saving.
    public class MultilingualFontDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, lower right row, lower right column, width, height
                Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 200, 100);

                // Set multilingual text (Latin + CJK characters)
                shape.Text = "English 中文 日本語";

                // Specify fonts for different script types
                shape.TextOptions.FarEastName = "MS Mincho"; // Font for Far East (CJK) characters
                shape.TextOptions.LatinName = "Arial";      // Font for Latin characters

                // Save the workbook
                workbook.Save("MultilingualFontDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            MultilingualFontDemo.Run();
        }
    }
}
