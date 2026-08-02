// Title: Set a PNG as worksheet background using a MemoryStream in Aspose.Cells for .NET (C#)
// Description: Loads a PNG file into a MemoryStream, converts it to a byte array, assigns it to the Worksheet.BackgroundImage property, and saves the workbook. The example demonstrates an in‑memory approach that avoids temporary files when adding a background image to an Excel sheet with Aspose.Cells.
// Keywords: Aspose.Cells | C# | .NET | MemoryStream | worksheet background image | PNG to Excel | in‑memory image | byte array background | Excel workbook example | code snippet
// Common Searches: Aspose.Cells set worksheet background from PNG | C# MemoryStream background image Excel | How to add background image to Excel sheet using Aspose.Cells | Load PNG into byte array for Aspose.Cells | Create Excel file with background image without saving temp file
// Developer Intent: Add a PNG image as the worksheet background by reading it into a MemoryStream and applying it through Aspose.Cells.
// Use Cases: Brand a generated report with a company logo as a repeated background. | Apply a watermark to a workbook template without writing the image to disk. | Dynamically change worksheet backgrounds based on user‑selected images stored in memory.
// AI Prompts: Show a C# method that takes a PNG path, loads it into a MemoryStream, and sets it as the worksheet background using Aspose.Cells. | Explain how to replace an existing worksheet background with a new image from a byte array in Aspose.Cells for .NET. | Provide error‑handling code for missing PNG files when assigning a background image to an Excel worksheet.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBackgroundImageDemo
{
    // Loads a PNG file into a MemoryStream, converts it to a byte array, assigns it to the Worksheet.BackgroundImage property, and saves the workbook. The example demonstrates an in‑memory approach that avoids temporary files when adding a background image to an Excel sheet with Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // Path to the PNG image file
            string pngFilePath = "background.png";

            // Verify that the image file exists
            if (!File.Exists(pngFilePath))
            {
                Console.WriteLine($"Image file not found: {pngFilePath}");
                return;
            }

            // Create a MemoryStream from the PNG file
            using (MemoryStream imageStream = new MemoryStream(File.ReadAllBytes(pngFilePath)))
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Assign the image bytes as the worksheet background image
                worksheet.BackgroundImage = imageStream.ToArray();

                // Save the workbook (using Aspose.Cells provided save method)
                string outputPath = "WorkbookWithBackground.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved with background image: {Path.GetFullPath(outputPath)}");
            }
        }
    }
}
