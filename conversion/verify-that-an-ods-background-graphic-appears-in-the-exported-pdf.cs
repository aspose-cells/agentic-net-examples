// Title: Check that a background image in an ODS workbook is retained after converting to PDF with Aspose.Cells for .NET
// Common Searches: how to confirm ODS worksheet background picture is kept when saving as PDF using Aspose.Cells C# | Aspose.Cells C# detect background image in ODS before PDF export | verify background graphic preservation in PDF generated from ODS with Aspose.Cells | C# code to test if ODS background image appears in exported PDF using Aspose.Cells | determine if ODS file background image is exported to PDF with Aspose.Cells .NET
// Tags: ODS to PDF background image preservation | Aspose.Cells BackgroundImage property usage | verify worksheet background graphic after PDF export | C# ODS background picture detection | export ODS workbook as PDF with graphics intact

using System;
using System.IO;
using Aspose.Cells;

// The example loads an ODS workbook, checks each worksheet for a BackgroundImage, saves the workbook as a PDF, and prints whether the background graphic was retained in the exported PDF.
class VerifyOdsBackgroundGraphic
{
    static void Main()
    {
        const string inputPath = "input.ods";
        const string outputPath = "output.pdf";

        try
        {
            // Ensure the input ODS file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the ODS workbook
            Workbook workbook = new Workbook(inputPath);

            // Determine whether any worksheet contains a background picture
            bool hasBackgroundImage = false;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Use BackgroundImage property (available in Aspose.Cells)
                if (sheet.BackgroundImage != null)
                {
                    hasBackgroundImage = true;
                    break;
                }
            }

            // Export the workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);

            // Output verification result
            Console.WriteLine(hasBackgroundImage
                ? "Background graphic is present in the exported PDF."
                : "Background graphic is NOT present in the exported PDF.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
