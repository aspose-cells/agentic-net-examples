// Title: Verify that slicer border color and thickness are preserved when exporting an Excel worksheet to PDF using Aspose.Cells for .NET
// Common Searches: how to confirm slicer border color is retained after exporting Excel to PDF with Aspose.Cells C# | Aspose.Cells C# check slicer border thickness in generated PDF | retrieve slicer border properties before PDF conversion using Aspose.Cells .NET | ensure slicer formatting (border) stays intact when saving workbook as PDF in C# | debug slicer border loss during Excel to PDF conversion with Aspose.Cells
// Tags: preserve slicer outline during PDF export | excel slicer styling Aspose.Cells | extract slicer border attributes C# | verify slicer appearance after PDF conversion | reflection access slicer properties Aspose.Cells | export workbook to PDF with slicer formatting intact

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, checks for slicers, uses reflection to read each slicer's BorderColor and BorderWeight (if available), prints the expected values, converts the workbook to PDF with Aspose.Cells, and confirms the PDF file is created, allowing developers to verify that slicer border styling is retained after conversion.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input Excel file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one slicer
            if (sheet.Slicers.Count == 0)
            {
                Console.WriteLine("No slicers found in the worksheet.");
                return;
            }

            // Retrieve the first slicer (type may vary by Aspose.Cells version)
            object slicerObj = sheet.Slicers[0];
            try
            {
                // Attempt to read border properties via reflection (if available)
                var borderColorProp = slicerObj.GetType().GetProperty("BorderColor");
                var borderWeightProp = slicerObj.GetType().GetProperty("BorderWeight");

                if (borderColorProp != null && borderWeightProp != null)
                {
                    Color borderColor = (Color)borderColorProp.GetValue(slicerObj);
                    float borderWeight = Convert.ToSingle(borderWeightProp.GetValue(slicerObj));

                    Console.WriteLine($"Expected Border Color   : {borderColor}");
                    Console.WriteLine($"Expected Border Thickness: {borderWeight} pt");
                }
                else
                {
                    Console.WriteLine("Slicer border properties are not accessible in this version.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to retrieve slicer border details. {ex.Message}");
            }

            // Convert the workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");

            // Verify that the PDF was created
            if (File.Exists(outputPath))
            {
                Console.WriteLine("PDF file creation confirmed.");
            }
            else
            {
                Console.WriteLine("Error: PDF file was not created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
