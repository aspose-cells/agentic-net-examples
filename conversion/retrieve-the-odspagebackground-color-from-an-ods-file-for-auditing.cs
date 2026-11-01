// Title: Read ODS page background color using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get default workbook background color from an ODS file | How to extract ODS page background ARGB values with Aspose.Cells .NET | Retrieve ODS sheet background color for auditing using Aspose.Cells | Aspose.Cells read ODS default style background color in C# | Get ODS workbook background color programmatically with Aspose.Cells
// Tags: Aspose.Cells read ODS default background color | C# extract ODS page background ARGB | Aspose.Cells workbook default style color .NET | ODS background color auditing with Aspose | retrieve ODS page background using Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example verifies that an input.ods file exists, loads it into an Aspose.Cells Workbook, accesses the first worksheet, reads the workbook's default style background color (which represents the ODS page background), and prints the ARGB components while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.ods";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the ODS file
            var workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            var worksheet = workbook.Worksheets[0];

            // Retrieve the workbook's default background color (as worksheet background is not directly exposed)
            Color backgroundColor = workbook.DefaultStyle.BackgroundColor;

            // Output the ARGB components of the color
            Console.WriteLine($"Workbook default background color - A:{backgroundColor.A}, R:{backgroundColor.R}, G:{backgroundColor.G}, B:{backgroundColor.B}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
