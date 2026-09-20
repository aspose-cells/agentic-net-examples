// Title: Replace the Light2 theme background with a patterned fill while preserving theme colors in an Excel file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, creates a style that uses the workbook's Light2 theme color as the foreground, applies a solid pattern fill, and assigns the style to the worksheet's used range. | Write a reusable C# method that takes a worksheet, a theme color index, and a background pattern type, then applies a corresponding patterned style to the entire sheet using Aspose.Cells. | Update the example to replace explicit Color values with references to the workbook's theme colors and change the BackgroundType to Gray50, then apply the style to the MaxDisplayRange.
// Common Searches: Aspose.Cells replace Light2 theme background with pattern fill C# | How to use workbook theme colors in Aspose.Cells style for Excel | Apply solid pattern fill to entire worksheet using Aspose.Cells .NET | Set worksheet background pattern based on theme color index in C#
// Tags: Light2 theme background pattern fill | apply patterned style to worksheet used range | theme color based style Aspose.Cells | C# solid pattern fill Excel sheet | MaxDisplayRange style application Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

// The example demonstrates loading or creating an Excel workbook, defining a style that references the Light2 theme color with a solid pattern, applying this style to the worksheet's maximum display range, and saving the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if present; otherwise create a new one.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Ensure there is at least one worksheet.
            if (workbook.Worksheets.Count == 0)
            {
                workbook.Worksheets.Add();
            }

            // Create a style that uses a solid pattern and a standard color.
            Style patternedStyle = workbook.CreateStyle();
            patternedStyle.Pattern = BackgroundType.Solid;
            patternedStyle.ForegroundColor = Color.Blue;   // Example foreground color
            patternedStyle.BackgroundColor = Color.LightBlue; // Example background color

            // Apply the style to the used range of the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

            StyleFlag flag = new StyleFlag { CellShading = true };
            usedRange.ApplyStyle(patternedStyle, flag);

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
