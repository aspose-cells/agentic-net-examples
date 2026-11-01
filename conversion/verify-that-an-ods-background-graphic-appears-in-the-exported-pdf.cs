// Title: Check if an ODS worksheet background image is preserved when exporting to PDF with Aspose.Cells for .NET
// Common Searches: C# code to verify that a background image in an ODS file is kept after saving as PDF with Aspose.Cells | how to detect if an ODS worksheet background graphic is exported to PDF using Aspose.Cells .NET | sample program to check ODS background picture presence before and after PDF conversion in C#
// Tags: ODS to PDF background image preservation | worksheet background graphic detection Aspose.Cells | export workbook with background picture to PDF | validate background image after ODS conversion | Aspose.Cells background image check C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an ODS workbook, scans each worksheet for a BackgroundImage, records whether a background graphic exists, saves the workbook as a PDF, and writes the detection result to the console.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.ods";
            const string outputPath = "output.pdf";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the ODS workbook
            Workbook workbook = new Workbook(inputPath);

            // Check for background graphics in any worksheet
            bool backgroundGraphicFound = false;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.BackgroundImage != null)
                {
                    backgroundGraphicFound = true;
                    break;
                }
            }

            // Export the workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);

            // Output verification result
            Console.WriteLine("Background graphic present: " + backgroundGraphicFound);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
