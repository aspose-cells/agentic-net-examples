// Title: C# example: load an XLSX workbook, apply a custom style to a range, and export to PDF with Aspose.Cells
// Common Searches: how to load an existing .xlsx file and convert it to PDF with custom cell formatting using Aspose.Cells in C# | Aspose.Cells C# apply background color and font style to a range before saving as PDF | saving a styled Excel worksheet as PDF with Aspose.Cells .NET API | C# code to create a custom style and apply it to cells A1:C3 in Aspose.Cells | export Excel to PDF while preserving cell alignment and colors using Aspose.Cells
// Tags: load xlsx workbook Aspose.Cells C# | create custom cell style Aspose.Cells | apply style to cell range Aspose.Cells | export styled worksheet to PDF Aspose.Cells | save workbook as PDF C# Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The program loads an XLSX workbook (or creates a new one if missing), defines a custom style with Arial font, blue text, light‑yellow background, and centered alignment, applies this style to the A1:C3 range on the first worksheet, and then saves the workbook as a PDF document using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
                // Optional: add placeholder data
                Worksheet newSheet = workbook.Worksheets[0];
                newSheet.Cells["A1"].PutValue("Placeholder");
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a custom style
            Style customStyle = workbook.CreateStyle();
            customStyle.Font.Name = "Arial";
            customStyle.Font.Size = 12;
            customStyle.Font.Color = Color.Blue;
            customStyle.ForegroundColor = Color.LightYellow;
            customStyle.Pattern = BackgroundType.Solid;
            customStyle.HorizontalAlignment = TextAlignmentType.Center;
            customStyle.VerticalAlignment = TextAlignmentType.Center;

            // Apply the custom style to a specific range (e.g., A1:C3)
            AsposeRange targetRange = sheet.Cells.CreateRange("A1:C3");
            StyleFlag styleFlag = new StyleFlag { All = true };
            targetRange.ApplyStyle(customStyle, styleFlag);

            // Save the workbook as a PDF document
            workbook.Save(outputPath, SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
