// Title: Add a signature line shape to cell B2 and apply password protection to the first worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Insert a SignatureLine shape at row 2, column 2, set its width to 200 points and height to 50 points, then protect the worksheet with the password "myPassword" using Aspose.Cells in C#. | Create a signature line object, place it on the worksheet at B2, resize it, enable full worksheet protection with a password, and save the workbook. | Load an existing Excel file, add a signature line shape to cell B2, adjust its dimensions, lock the sheet with a password, and write the result to a new file using Aspose.Cells for .NET.
// Common Searches: how to insert a signature line at a specific cell with Aspose.Cells C# | Aspose.Cells protect worksheet after adding a shape with password | set signature line dimensions in Excel using Aspose.Cells .NET | C# code to add a signature line and lock the sheet in Aspose.Cells | Aspose.Cells example for adding a signature line and applying worksheet protection
// Tags: signature line insertion Aspose.Cells | shape dimension setting Aspose.Cells | worksheet password protection Aspose.Cells | protect Excel sheet C# Aspose.Cells | add signature line to B2 cell Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an existing workbook, adds a SignatureLine shape at cell B2 with a width of 200 points and height of 50 points, applies full worksheet protection using the password "myPassword", and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a SignatureLine object with desired size
            SignatureLine signatureLine = new SignatureLine();
            // The size is set on the returned Shape after adding to the sheet
            // Add the signature line shape at row 2, column 2 (zero‑based)
            Shape shape = sheet.Shapes.AddSignatureLine(1, 1, signatureLine);
            shape.Width = 200;   // width in points
            shape.Height = 50;   // height in points

            // Protect the worksheet with a password (old password is empty)
            sheet.Protect(ProtectionType.All, "myPassword", string.Empty);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
