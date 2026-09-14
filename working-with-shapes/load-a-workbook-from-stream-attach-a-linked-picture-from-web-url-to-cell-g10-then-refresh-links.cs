// Title: Load an Excel workbook from a stream and add a linked picture from a web URL to cell G10 using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads an Excel file from a Stream, uses Worksheet.Pictures.Add to attach a web‑linked image to cell G10, and saves the workbook as an XLSX file with Aspose.Cells. | Show how to insert a linked picture from an external URL into a specific cell after loading a workbook from memory using Aspose.Cells for .NET. | Provide an example that loads a workbook from a stream, adds a URL‑based picture to cell G10, and writes the updated file to disk with Aspose.Cells.
// Common Searches: c# aspose.cells add linked picture from URL to specific cell | load excel workbook from memory stream and insert image into G10 using Aspose.Cells | worksheet pictures add method with external URL example asp.net | how to save workbook after adding a web‑linked picture with Aspose.Cells | asp.net refresh linked pictures after adding them to worksheet
// Tags: load workbook from stream Aspose.Cells | add linked picture to worksheet cell Aspose.Cells | Worksheet.Pictures.Add URL overload | save workbook as xlsx Aspose.Cells | refresh external links Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel file from a Stream, inserts a linked picture from a web URL into cell G10 of the first worksheet, and saves the modified workbook as an XLSX file using Aspose.Cells for .NET.
public class Program
{
    public static void Main()
    {
        try
        {
            // Obtain the input workbook stream safely
            using (Stream inputStream = GetInputStream())
            {
                if (inputStream == null)
                {
                    Console.WriteLine("Input file not found. Operation aborted.");
                    return;
                }

                // Load workbook from the stream
                Workbook workbook = new Workbook(inputStream);

                // Access the first worksheet (index 0)
                Worksheet sheet = workbook.Worksheets[0];

                // URL of the picture to link
                string pictureUrl = "https://example.com/image.png";

                try
                {
                    // Add a linked picture to cell G10 (row 9, column 6 – zero‑based indices)
                    // This overload adds a picture from a URL and treats it as a linked picture.
                    int pictureIndex = sheet.Pictures.Add(9, 6, pictureUrl);
                    // Optionally retrieve the picture object if further manipulation is needed
                    // Picture picture = sheet.Pictures[pictureIndex];
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to add linked picture: {ex.Message}");
                }

                // Save the modified workbook
                string outputPath = "output.xlsx";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                using (FileStream outStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Save(outStream, SaveFormat.Xlsx);
                }

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Returns a FileStream for the input workbook if the file exists; otherwise null
    private static Stream GetInputStream()
    {
        const string inputPath = "input.xlsx";

        if (!File.Exists(inputPath))
        {
            return null;
        }

        return new FileStream(inputPath, FileMode.Open, FileAccess.Read);
    }
}
