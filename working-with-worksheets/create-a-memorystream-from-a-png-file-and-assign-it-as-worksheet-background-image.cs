// Title: Set Worksheet Background from PNG and Export Workbook via MemoryStream – Aspose.Cells for .NET
// Description: Demonstrates how to load a PNG file into a byte array, assign it to a worksheet's BackgroundImage property, and save the workbook directly to a MemoryStream with optional file output, using Aspose.Cells in C#.
// Keywords: Aspose.Cells set worksheet background image | load PNG into worksheet background .NET | Workbook.SaveToStream Aspose.Cells | C# MemoryStream workbook export | assign background image byte array Aspose.Cells | background image without temporary files
// Common Searches: How to add a PNG background to an Aspose.Cells worksheet | Save Aspose.Cells workbook to MemoryStream after setting background | C# load image file as byte[] for worksheet background | Export Aspose.Cells workbook directly to stream | Apply background image to first worksheet in code
// Developer Intent: Apply a PNG as the worksheet background and obtain the workbook as an in‑memory stream.
// Use Cases: Generate branded Excel reports with a background image and stream them to a web client without creating temporary files. | Create temporary workbooks in server‑side services, apply custom backgrounds, and keep the result in memory for further processing. | Batch process multiple PNG assets, set each as a worksheet background, and store the resulting workbooks in MemoryStream objects for downstream APIs.
// AI Prompts: Provide C# code using Aspose.Cells that reads a PNG file, sets it as the background of the first worksheet, and returns the workbook as a MemoryStream. | Show an example that checks for the PNG file, assigns it via the BackgroundImage property, and saves the workbook directly to a stream. | Write a method that accepts a Stream containing a PNG image and outputs a MemoryStream of an Aspose.Cells workbook with that image applied as the worksheet background.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to load a PNG file into a byte array, assign it to a worksheet's BackgroundImage property, and save the workbook directly to a MemoryStream with optional file output, using Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the PNG file
            string pngPath = "background.png";

            // Load the PNG file into a byte array if it exists (load rule)
            if (File.Exists(pngPath))
            {
                byte[] pngBytes = File.ReadAllBytes(pngPath);
                // Assign the background image using the byte array
                worksheet.BackgroundImage = pngBytes;
            }
            else
            {
                Console.WriteLine($"Warning: Background image file '{pngPath}' not found. Skipping background assignment.");
            }

            // Save the workbook to a MemoryStream (save rule)
            using (MemoryStream workbookStream = workbook.SaveToStream())
            {
                // Write the workbook stream to a file (optional)
                using (FileStream file = new FileStream("WorkbookWithBackground.xls", FileMode.Create, FileAccess.Write))
                {
                    workbookStream.WriteTo(file);
                }
            }

            Console.WriteLine("Workbook created successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
