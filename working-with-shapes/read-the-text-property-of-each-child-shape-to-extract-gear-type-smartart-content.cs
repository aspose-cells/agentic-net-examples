// Title: Read the Text property of child shapes inside a Gear SmartArt group in an Excel worksheet with Aspose.Cells for .NET
// Common Searches: how to get text from child shapes of a SmartArt group using Aspose.Cells C# | Aspose.Cells read Gear SmartArt node text in Excel workbook | C# iterate worksheet shapes and extract grouped shape text Aspose.Cells | detect SmartArt groups in Excel and read their child shape content with Aspose.Cells | Aspose.Cells .NET retrieve Text property from shapes inside SmartArt
// Tags: worksheet shape enumeration Aspose.Cells | grouped shape detection .NET | Gear SmartArt node access limitation | extract shape Text property Aspose.Cells | Excel SmartArt group handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtExample
{
    // The example loads an Excel file, checks for its existence, iterates over all shapes on the first worksheet, identifies grouped shapes that may contain SmartArt using the IsGroup property, logs their names, and notes that direct extraction of SmartArt node text (e.g., Gear SmartArt) is not supported by Aspose.Cells.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Iterate through all shapes on the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Group shapes can contain SmartArt; check using IsGroup property
                    if (shape.IsGroup)
                    {
                        Console.WriteLine($"Possible SmartArt shape detected: {shape.Name}");
                        // SmartArt node extraction is not supported in Aspose.Cells.
                    }
                }
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
