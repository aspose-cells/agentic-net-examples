// Title: Set 30% fill transparency for a WordArt shape in an Excel worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Create a new workbook, add a WordArt text effect shape, and set its Fill.Transparency to 0.3 using Aspose.Cells in C#. | Programmatically adjust the background opacity of a WordArt shape to 30% and save the worksheet as XLSX with Aspose.Cells for .NET. | Insert a WordArt shape into a worksheet and modify its fill color and transparency to achieve a subtle background effect using C#.
// Common Searches: Aspose.Cells C# set WordArt shape fill transparency to 30 percent | How to change opacity of a text effect shape in Excel using Aspose.Cells .NET | Example code for adjusting WordArt background opacity in an Excel file with Aspose.Cells | C# Aspose.Cells modify WordArt fill transparency before saving workbook
// Tags: Aspose.Cells set shape opacity | WordArt shape fill opacity C# | Excel worksheet WordArt opacity adjustment | C# modify shape fill opacity Aspose.Cells | text effect shape transparency .NET

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates creating a workbook, inserting a WordArt text effect shape, setting its fill transparency to 30%, and saving the file as XLSX using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a WordArt (text effect) shape to the worksheet using a preset effect
            Shape wordArt = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1, // preset text effect
                "Sample WordArt",                // text
                "Arial",                         // font name
                36,                              // font size
                false,                           // bold
                false,                           // italic
                2,                               // upper left row
                2,                               // upper left column
                0,                               // upper left pixel (top)
                0,                               // left pixel
                100,                             // height
                300                              // width
            );

            // Adjust the fill transparency to 30% (0.3)
            // Transparency range: 0 (opaque) to 1 (fully transparent)
            wordArt.Fill.Transparency = 0.3;

            // Optionally set a light fill color for a subtle background effect
            // Note: ForeColor property may not be available in some versions; this line can be omitted or adjusted as needed.
            // wordArt.Fill.ForeColor = Color.LightGray;

            // Save the workbook to a file
            workbook.Save("WordArtTransparency.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
