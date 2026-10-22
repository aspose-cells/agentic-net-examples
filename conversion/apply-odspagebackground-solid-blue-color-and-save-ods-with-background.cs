// Title: Set a solid blue background for an ODS worksheet using Aspose.Cells for .NET and save the file
// Common Searches: how to apply a solid blue background to an ODS sheet with Aspose.Cells in C# | Aspose.Cells C# set worksheet background color before saving as ODS | create ODS file with blue page background using Aspose.Cells for .NET | C# code to fill entire worksheet range with solid color and export to ODS
// Tags: solid background style Aspose.Cells | worksheet range fill color C# | export workbook to ODS format Aspose | backgroundtype solid usage Aspose.Cells | ODS generation with page color .NET

using Aspose.Cells;
using System;
using System.Drawing;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Cells.Range;

// The example creates a new Workbook, defines a solid blue style, applies it to a large worksheet range, and saves the result as an ODS file named 'BlueBackground.ods' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define a solid blue fill style
            Style blueStyle = workbook.CreateStyle();
            blueStyle.ForegroundColor = Color.Blue;
            blueStyle.Pattern = BackgroundType.Solid;

            // Apply the style to a large range covering typical sheet area
            AsposeRange range = sheet.Cells.CreateRange("A1", "Z1000");
            range.ApplyStyle(blueStyle, new StyleFlag { All = true });

            // Save the workbook as ODS (lifecycle save rule)
            string outputPath = "BlueBackground.ods";
            workbook.Save(outputPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
