// Title: Change the default hyperlink font color to a custom blue shade and save the workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change default hyperlink color to RGB 0,112,192 | How to set a custom blue color for all hyperlinks in an Excel workbook with Aspose.Cells .NET | Programmatically update hyperlink theme color and save the workbook using Aspose.Cells | C# Aspose.Cells modify default style to change hyperlink font color and underline | Save Excel file after changing default hyperlink style with Aspose.Cells
// Tags: modify default hyperlink style Aspose.Cells | set hyperlink font color RGB Aspose.Cells | apply single underline to hyperlink style C# | save workbook after style changes Aspose.Cells | custom blue hyperlink color Excel .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, accesses the workbook's default style, changes the hyperlink font color to the custom blue RGB value (0,112,192), ensures the hyperlinks remain underlined, reassigns the updated style, and saves the workbook to a new file while handling potential errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the default style (includes hyperlink formatting)
            Style hyperlinkStyle = workbook.DefaultStyle;

            // Set a new shade of blue for the hyperlink font color (RGB 0,112,192)
            hyperlinkStyle.Font.Color = Color.FromArgb(0, 112, 192);

            // Ensure underline is applied for hyperlinks
            hyperlinkStyle.Font.Underline = FontUnderlineType.Single;

            // Apply the modified style back to the workbook
            workbook.DefaultStyle = hyperlinkStyle;

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
