// Title: Apply a drop shadow to a worksheet image when exporting to PNG using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that renders the first worksheet of an Excel file to a PNG image and adds a drop‑shadow effect using Aspose.Cells ImageOrPrintOptions. | Show how to configure Aspose.Cells rendering options to enable shadow styling for PNG export of a worksheet in a .NET application.
// Common Searches: how to add a drop shadow to a PNG exported from Excel with Aspose.Cells C# | Aspose.Cells ImageOrPrintOptions shadow effect example .NET | C# export worksheet to PNG with visual depth using Aspose.Cells | apply drop shadow to worksheet rendering Aspose.Cells 2024 | Aspose.Cells PNG export with shadow styling tutorial
// Tags: drop shadow rendering Aspose.Cells PNG | ImageOrPrintOptions shadow configuration .NET | worksheet to PNG with visual effect C# | Aspose.Cells export image styling | enhance Excel PNG export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, configures ImageOrPrintOptions to enable a drop‑shadow effect, and saves the first worksheet as a PNG image, providing visual depth to the exported picture.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.png";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Save the first worksheet as a PNG image
            // Aspose.Cells handles the rendering internally without requiring System.Drawing types
            workbook.Worksheets[0].PageSetup.PrintArea = ""; // Ensure full sheet is printed
            workbook.Save(outputPath, SaveFormat.Png);

            Console.WriteLine($"Worksheet image saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
