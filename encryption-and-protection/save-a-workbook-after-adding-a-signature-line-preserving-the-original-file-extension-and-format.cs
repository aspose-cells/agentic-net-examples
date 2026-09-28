// Title: Add a signature line to an existing Excel workbook and save it with the original file extension using Aspose.Cells for .NET
// AI Prompts: Load any Excel file with Aspose.Cells, insert a SignatureLine shape at cell E5, and save the workbook using a '_Signed' suffix while keeping the original format. | Create a C# console app that opens a workbook, adds a digital signature line to the first worksheet, and lets Aspose.Cells infer the save format from the file extension. | Write code that reads an Excel workbook, places a SignatureLine object on the sheet, and writes the modified file back preserving its original extension.
// Common Searches: how to insert a signature line into an existing Excel workbook using Aspose.Cells C# | save modified Excel file with original extension Aspose.Cells .NET | preserve original workbook format when saving after adding a signature line | Aspose.Cells add SignatureLine shape to worksheet and keep file type
// Tags: add signature line Aspose.Cells C# | preserve workbook format Aspose.Cells | save workbook with original extension Aspose.Cells | signature line shape Excel worksheet | load and modify Excel file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads an existing workbook, adds a SignatureLine shape at cell E5 on the first worksheet, and saves the file with a "_Signed" suffix while preserving the original file extension and format.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the original workbook (any supported format)
            string inputPath = @"C:\Docs\SampleWorkbook.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (preserves original format)
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a SignatureLine object and set its supported properties
            SignatureLine sigLine = new SignatureLine
            {
                Signer = "John Doe",
                Instructions = "Please sign here"
                // Note: SignerTitle and SignerEmail are not available in this API version
            };

            // Add the signature line to the worksheet at cell E5 (row 4, column 4)
            Shape shape = sheet.Shapes.AddSignatureLine(4, 4, sigLine);

            // Build the output file name while preserving the original extension
            string directory = Path.GetDirectoryName(inputPath) ?? string.Empty;
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
            string extension = Path.GetExtension(inputPath);
            string outputPath = Path.Combine(directory, $"{fileNameWithoutExt}_Signed{extension}");

            // Ensure the output directory exists
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook; Aspose.Cells determines the format from the file extension
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
