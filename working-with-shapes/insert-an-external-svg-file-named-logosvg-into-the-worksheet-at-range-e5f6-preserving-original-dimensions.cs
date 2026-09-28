// Title: Insert an external SVG file into cells E5:F6 without scaling using Aspose.Cells for .NET
// AI Prompts: Add the SVG file 'logo.svg' to the worksheet at cells E5:F6, keep its original size, and save the workbook as output.xlsx with Aspose.Cells C# API. | Generate C# code that verifies logo.svg exists, inserts it using the Pictures.Add overload that preserves dimensions, and exports the file to XLSX.
// Common Searches: Aspose.Cells C# add external SVG to specific cells without resizing | How to place an SVG image in Excel range E5:F6 using Aspose.Cells | Preserve original SVG dimensions when inserting picture with Aspose.Cells .NET | C# Pictures.Add overload example for SVG file in Aspose.Cells | Insert external vector graphic into Excel worksheet using Aspose.Cells API
// Tags: Aspose.Cells insert SVG picture | Pictures.Add overload preserve dimensions | C# add external SVG to worksheet | Insert image into specific cell range Aspose.Cells | Save workbook with SVG image .NET

using System;
using System.IO;
using Aspose.Cells;

// The program creates a new workbook, checks for the presence of logo.svg, inserts the SVG into cells E5:F6 using the Pictures.Add overload that retains the original image size, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the external SVG file
            string svgPath = "logo.svg";

            // Insert the SVG picture if the file exists
            if (File.Exists(svgPath))
            {
                // Insert the external SVG file at cell E5 (row index 4, column index 4)
                // Using the overload without lower‑right cell preserves the original SVG dimensions.
                sheet.Pictures.Add(4, 4, svgPath);
            }
            else
            {
                Console.WriteLine($"Warning: SVG file '{svgPath}' not found. Skipping picture insertion.");
            }

            // Save the workbook (lifecycle rule)
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
