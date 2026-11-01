// Title: Insert a PNG image into a worksheet cell and attach a file:// hyperlink that opens a PDF when clicked using Aspose.Cells for .NET (C#)
// AI Prompts: Insert a picture from a PNG file into cell B2 of a new workbook and set its Hyperlink.Address to a local PDF file path with a screen tip. | Add an image to an Excel worksheet, assign a file:// hyperlink that opens a document, and save the workbook using the Aspose.Cells C# API. | Create a workbook, embed a PNG picture, configure the picture's Hyperlink properties (Address and ScreenTip), and export the file as XLSX.
// Common Searches: Aspose.Cells C# how to add a picture to a specific cell and link it to a local PDF file | set hyperlink on an inserted image in Excel using Aspose.Cells .NET | embed PNG in worksheet and open PDF on click with Aspose.Cells | assign screen tip to picture hyperlink in Aspose.Cells C# example
// Tags: insert picture into worksheet cell Aspose.Cells | picture hyperlink to local file C# | set picture screen tip Aspose.Cells | save workbook with linked image .NET | add PNG image to Excel using Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The sample creates a new workbook, inserts a PNG image into cell B2, assigns a file:// hyperlink pointing to a PDF with a screen tip, and saves the workbook as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the image you want to insert
            string imagePath = "sample.png";

            // Ensure the image file exists before attempting to insert
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Insert the picture at cell B2 (row index 1, column index 1)
            int pictureIndex = worksheet.Pictures.Add(1, 1, imagePath);

            // Retrieve the inserted picture object
            Picture picture = worksheet.Pictures[pictureIndex];

            // Set hyperlink properties (the Hyperlink object is read‑only, so modify its members)
            picture.Hyperlink.Address = "file:///C:/Docs/target.pdf";
            picture.Hyperlink.ScreenTip = "Open target PDF";

            // Save the workbook to a file
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
