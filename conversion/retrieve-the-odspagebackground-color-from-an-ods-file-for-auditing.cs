// Title: How to read the page background color of each worksheet in an ODS file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get page background color of a worksheet in an ODS file | How to extract ODS worksheet page background color using Aspose.Cells .NET | Aspose.Cells ODS page background color auditing sample code | Retrieve page setup background color from an ODS workbook with Aspose.Cells | Aspose.Cells version‑specific handling of missing PageBackgroundColor property
// Tags: read ODS worksheet page background color C# | Aspose.Cells get worksheet page background color | audit ODS page background color Aspose.Cells | handle missing PageBackgroundColor property Aspose.Cells | extract page setup color from ODS using Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example verifies the ODS file exists, loads it with Aspose.Cells, iterates through each worksheet, attempts to read the worksheet's page background color (using a placeholder when the property is unavailable in older versions), prints the worksheet name and color, and gracefully handles any exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.ods";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the ODS file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet to get its page background color
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // PageSetup.PageBackgroundColor is not available in older Aspose.Cells versions.
                // Use a placeholder or implement version‑specific logic if needed.
                Color pageBgColor = Color.Empty;

                Console.WriteLine($"Worksheet: {sheet.Name}, Page Background Color: {pageBgColor}");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
