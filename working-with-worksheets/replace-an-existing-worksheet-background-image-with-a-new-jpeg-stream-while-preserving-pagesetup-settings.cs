// Title: Replace Worksheet Background Image with a JPEG Stream in Aspose.Cells for .NET (Preserve Page Setup)
// Description: Loads an existing workbook, reads a JPEG file into a byte array, assigns the array to the worksheet's BackgroundImage property, and saves the file. The operation swaps the current background without altering any page‑setup settings such as margins, orientation, or print area.
// Keywords: Aspose.Cells replace worksheet background | set worksheet background from byte array | JPEG background image .NET | preserve page setup Aspose.Cells | update Excel background image programmatically
// Common Searches: how to change worksheet background image Aspose.Cells C# | replace Excel sheet background without affecting margins | set background image from stream Aspose.Cells | update worksheet background while keeping page layout
// Developer Intent: Swap the existing worksheet background with a new JPEG byte stream while keeping all page‑setup configurations unchanged.
// Use Cases: Apply a corporate logo as a background to a generated report without modifying print layout. | Replace a template placeholder image with a user‑provided picture before distribution, retaining margins and orientation. | Refresh a worksheet’s background after an upload, ensuring print settings remain intact.
// AI Prompts: Write C# code that loads a workbook, replaces a worksheet's background using a JPEG byte array, and saves the file while preserving page‑setup settings. | Show an Aspose.Cells example that sets BackgroundImage from a MemoryStream and verifies that margins and orientation are unchanged. | Explain how to test that page‑setup properties (orientation, margins, print area) stay the same after updating the worksheet background.

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // Loads an existing workbook, reads a JPEG file into a byte array, assigns the array to the worksheet's BackgroundImage property, and saves the file. The operation swaps the current background without altering any page‑setup settings such as margins, orientation, or print area.
    class ReplaceWorksheetBackground
    {
        static void Main()
        {
            try
            {
                // Path to the source workbook
                string sourceFile = "input.xlsx";
                if (!File.Exists(sourceFile))
                {
                    Console.WriteLine($"Source workbook not found: {sourceFile}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(sourceFile);
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the new background image
                string backgroundPath = "newBackground.jpg";
                if (File.Exists(backgroundPath))
                {
                    // Load image bytes and set as worksheet background
                    byte[] newBackground = File.ReadAllBytes(backgroundPath);
                    worksheet.BackgroundImage = newBackground;
                }
                else
                {
                    Console.WriteLine($"Background image not found: {backgroundPath}");
                }

                // Save the updated workbook
                string outputFile = "output.xlsx";
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved to {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
