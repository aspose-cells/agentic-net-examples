// Title: C# – Save Aspose.Cells Workbook as PNG with a Timestamped Filename in a Specified Output Folder
// Description: Shows how to ensure an "output" directory exists, build a filename that embeds the current date‑time (yyyyMMdd_HHmmss), set ImageSaveOptions for PNG, enable CreateDirectory, and write the workbook as a PNG image to the generated path.
// Keywords: Aspose.Cells PNG export C# | ImageSaveOptions CreateDirectory | timestamped filename C# | save workbook as image | generate output folder C# | Aspose.Cells image save options | date time filename PNG
// Common Searches: Aspose.Cells save workbook as PNG C# | C# export Excel to PNG with timestamp | ImageSaveOptions CreateDirectory example | unique PNG filename Aspose.Cells | save chart as PNG using Aspose.Cells C#
// Developer Intent: Export a workbook (or its chart) to a PNG file and store it in a chosen directory using a unique timestamped name.
// Use Cases: Automatically archive daily workbook snapshots as PNG files for audit trails. | Generate timestamped chart images for scheduled email or dashboard reports. | Maintain versioned visual copies of workbooks in a shared folder for collaborative review.
// AI Prompts: Write C# code that uses Aspose.Cells to export a workbook to PNG, creates the output folder if needed, and names the file with the current timestamp. | Show how to set ImageSaveOptions.CreateDirectory = true and save a workbook as a PNG with a date‑time based filename. | Explain how to customize the timestamp format in the PNG filename when saving an Aspose.Cells workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsImageSaveDemo
{
    // Shows how to ensure an "output" directory exists, build a filename that embeds the current date‑time (yyyyMMdd_HHmmss), set ImageSaveOptions for PNG, enable CreateDirectory, and write the workbook as a PNG image to the generated path.
    class Program
    {
        static void Main()
        {
            // Define the output directory
            string outputDir = "output";

            // Ensure the directory exists (CreateDirectory will also be set in save options)
            Directory.CreateDirectory(outputDir);

            // Generate a timestamped filename
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string fileName = $"Workbook_{timestamp}.png";
            string outputPath = Path.Combine(outputDir, fileName);

            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells PNG Export");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B1"].PutValue("Sample");
            sheet.Cells["B2"].PutValue(456);

            // Configure image save options for PNG format
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png);
            // Automatically create the directory if it does not exist
            saveOptions.CreateDirectory = true;

            // Save the workbook as a PNG image with the timestamped filename
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved as PNG image to: {outputPath}");
        }
    }
}
