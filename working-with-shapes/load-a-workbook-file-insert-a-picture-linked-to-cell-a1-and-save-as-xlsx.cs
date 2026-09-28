// Title: Load an existing XLSX workbook, insert a PNG picture into cell A1 with MoveAndSize placement, and save the file using Aspose.Cells for .NET
// AI Prompts: Load a workbook from a file, add a PNG image to cell A1 of the first worksheet, set the picture's Placement to MoveAndSize, and save the result as a new XLSX file. | Insert a linked picture into the top‑left cell of an Excel sheet and ensure it resizes with rows and columns using Aspose.Cells in C#. | Check that the source workbook and image exist, embed the image at A1, adjust its placement, and export the modified workbook to XLSX with Aspose.Cells.
// Common Searches: Aspose.Cells C# add image to specific cell and keep it linked | how to set picture placement MoveAndSize when inserting picture with Aspose.Cells | save workbook after embedding PNG picture in cell A1 using Aspose.Cells for .NET | load existing Excel file and insert picture at A1 with Aspose.Cells API | C# Aspose.Cells insert picture and preserve size on row column changes
// Tags: insert picture into cell A1 Aspose.Cells | picture placement MoveAndSize C# | load workbook and embed PNG Aspose.Cells | save workbook as XLSX with image Aspose.Cells | linked image in Excel cell .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads input.xlsx, verifies image.png exists, inserts the PNG into cell A1 of the first worksheet with MoveAndSize placement, and saves the updated workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string imagePath = "image.png";
            const string outputPath = "output.xlsx";

            // Ensure required files exist
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add picture to cell A1 (row 0, column 0) and set its placement
            int pictureIndex = sheet.Pictures.Add(0, 0, imagePath);
            Picture picture = sheet.Pictures[pictureIndex];
            picture.Placement = PlacementType.MoveAndSize;

            // Save the workbook as XLSX
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
