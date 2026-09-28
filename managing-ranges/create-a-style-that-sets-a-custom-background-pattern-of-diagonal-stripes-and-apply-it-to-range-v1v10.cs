// Title: How to create a diagonal stripe background style and apply it to cells V1:V10 with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells that creates a Style with BackgroundType.DiagonalStripe, sets LightBlue foreground and White background colors, and applies it to the range V1:V10. | Show how to use a StyleFlag to apply only the cell shading (pattern) part of a custom style to a specific column range in an Aspose.Cells workbook.
// Common Searches: Aspose.Cells C# apply diagonal stripe pattern to a column range | How to set BackgroundType.DiagonalStripe for cells V1 to V10 in .NET | Using StyleFlag to apply only shading in Aspose.Cells workbook | Create custom cell style with foreground and background colors in Aspose.Cells | Apply pattern style to a range without affecting other cell properties Aspose.Cells
// Tags: Aspose.Cells create diagonal stripe style | apply style to range V1 V10 Aspose.Cells | StyleFlag cell shading only Aspose.Cells | BackgroundType.DiagonalStripe C# example | custom cell background pattern Aspose.Cells

using System;
using Aspose.Cells;
using System.Drawing;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Cells.Range;

// The example creates a new workbook, defines a Style with a diagonal stripe pattern (LightBlue foreground on White background), uses a StyleFlag to apply only the shading, applies the style to cells V1 through V10, and saves the file as StyledRange.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Create a new style.
            Style diagonalStyle = workbook.CreateStyle();

            // Set the background pattern to diagonal stripes.
            diagonalStyle.Pattern = BackgroundType.DiagonalStripe;

            // Define foreground and background colors for the pattern.
            diagonalStyle.ForegroundColor = Color.LightBlue;   // Color of the stripes.
            diagonalStyle.BackgroundColor = Color.White;      // Base color.

            // Prepare a StyleFlag to apply only the shading (pattern) part of the style.
            StyleFlag flag = new StyleFlag
            {
                CellShading = true   // Apply pattern and colors.
            };

            // Create the target range V1:V10.
            AsposeRange targetRange = sheet.Cells.CreateRange("V1", "V10");

            // Apply the style to the range.
            targetRange.ApplyStyle(diagonalStyle, flag);

            // Save the workbook to a file.
            string outputPath = "StyledRange.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
