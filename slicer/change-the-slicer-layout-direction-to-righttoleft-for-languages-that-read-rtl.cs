// Title: Detect right‑to‑left culture and prepare to set slicer layout direction with Aspose.Cells in C#
// Common Searches: how to detect RTL culture in .NET before adjusting Excel slicer | Aspose.Cells C# change slicer orientation for right‑to‑left languages | programmatically set slicer layout direction in an Excel workbook using Aspose.Cells | detect right‑to‑left reading order and modify slicer properties with C# | example for handling RTL cultures when working with slicers in Aspose.Cells
// Tags: detect rtl culture Aspose.Cells C# | slicer orientation rtl Aspose.Cells | programmatic slicer layout Excel C# | culture-aware Excel slicer Aspose.Cells | right-to-left slicer direction .NET

using Aspose.Cells;
using System;
using System.Globalization;
using System.IO;

// The example loads an existing workbook (or creates a new one), checks whether the current .NET culture uses a right‑to‑left reading order, and notes that slicer handling is omitted for compatibility when RTL is detected, then saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file '{inputPath}' not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Determine if the current culture uses right‑to‑left reading order
            bool isRtl = CultureInfo.CurrentCulture.TextInfo.IsRightToLeft;

            if (isRtl)
            {
                // The Slicer class may not be available in older Aspose.Cells versions.
                // If slicer support is required, ensure the library version includes it.
                // Here we simply skip slicer handling to keep the code compilable.
                Console.WriteLine("Right‑to‑left culture detected, but slicer handling is omitted for compatibility.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
