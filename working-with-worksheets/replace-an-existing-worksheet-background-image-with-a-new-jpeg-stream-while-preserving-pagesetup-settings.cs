// Title: Replace Worksheet Background Image with JPEG Stream in Aspose.Cells for .NET (Preserve Page Setup)
// Description: Load an existing workbook, read a JPEG file into a byte array, assign it to Worksheet.BackgroundImage, and save the file while keeping all PageSetup settings unchanged.
// Keywords: Aspose.Cells background image | C# replace worksheet background | Worksheet.BackgroundImage byte array | preserve page setup Aspose.Cells | load JPEG as worksheet background
// Common Searches: Aspose.Cells change worksheet background without affecting page setup | C# set worksheet background from JPEG stream | replace Excel sheet background image Aspose | keep margins and orientation when updating worksheet background
// Developer Intent: Swap the current background picture of a worksheet for a new JPEG supplied as a stream, ensuring that page‑setup properties such as margins, orientation, and scaling remain intact.
// Use Cases: Refresh the visual theme of a template workbook while retaining its print layout. | Add a corporate logo as a background to generated reports without modifying existing print settings. | Rotate seasonal artwork in an existing spreadsheet without disturbing predefined page configurations.
// AI Prompts: Write C# code using Aspose.Cells to replace a worksheet's background image from a MemoryStream while preserving all PageSetup options. | Show an example that loads a JPEG file into a byte array, sets Worksheet.BackgroundImage, and includes robust error handling for missing files. | Explain how to verify that margins, orientation, and scaling stay the same after changing the worksheet background with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Load an existing workbook, read a JPEG file into a byte array, assign it to Worksheet.BackgroundImage, and save the file while keeping all PageSetup settings unchanged.
class ReplaceWorksheetBackground
{
    static void Main()
    {
        // Paths
        string workbookPath = "input.xlsx";
        string newImagePath = "newBackground.jpg";
        string outputPath = "output.xlsx";

        try
        {
            // Verify input files exist
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"Workbook file not found: {workbookPath}");

            if (!File.Exists(newImagePath))
                throw new FileNotFoundException($"Background image file not found: {newImagePath}");

            // Load the existing workbook (preserves all existing settings, including PageSetup)
            Workbook workbook = new Workbook(workbookPath);

            // Get the worksheet whose background image you want to replace
            Worksheet worksheet = workbook.Worksheets[0]; // adjust index if needed

            // Read the new JPEG image into a byte array
            byte[] newImageData = File.ReadAllBytes(newImagePath);

            // Replace the worksheet's background image while leaving PageSetup untouched
            worksheet.BackgroundImage = newImageData;

            // Save the workbook with the updated background image
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"File error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
