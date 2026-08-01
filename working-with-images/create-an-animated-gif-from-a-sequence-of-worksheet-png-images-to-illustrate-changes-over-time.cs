// Title: Create an animated GIF from Excel worksheets with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, then saves each worksheet or printed page as a frame in an animated GIF using Aspose.Cells SaveFormat.Gif. The sample checks the source file, handles errors, and outputs a GIF that visualizes workbook changes.
// Keywords: Aspose.Cells | C# | animated GIF | SaveFormat.Gif | export Excel to GIF | worksheet to GIF | Excel animation | convert Excel to GIF | Aspose.Cells GIF animation
// Common Searches: Aspose.Cells create animated GIF from workbook | C# export Excel sheets to GIF animation | SaveFormat.Gif multiple worksheets Aspose | how to generate GIF from Excel using Aspose.Cells | convert Excel worksheets to animated GIF .NET
// Developer Intent: Generate a GIF where each worksheet or printed page of an Excel file appears as a sequential frame.
// Use Cases: Show a time‑lapse of daily report sheets as a single GIF for quick review. | Create a compact preview of a multi‑sheet workbook to embed in emails or documentation. | Automate conversion of Excel dashboards into animated GIFs for web portals or presentations.
// AI Prompts: How can I set frame delay and loop count when saving a workbook as an animated GIF with Aspose.Cells? | Provide C# code to export only selected worksheets to an animated GIF using Aspose.Cells. | Explain how to control image quality, dimensions, and scaling of GIF frames generated from an Excel workbook.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, then saves each worksheet or printed page as a frame in an animated GIF using Aspose.Cells SaveFormat.Gif. The sample checks the source file, handles errors, and outputs a GIF that visualizes workbook changes.
class AnimatedGifFromWorksheets
{
    static void Main()
    {
        // Paths
        string inputExcel = "input.xlsx";   // source workbook
        string outputGif = "animation.gif"; // resulting animated GIF

        // Verify input file exists
        if (!File.Exists(inputExcel))
        {
            Console.WriteLine($"Input file '{inputExcel}' not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputExcel);

            // Save the workbook as an animated GIF.
            // Each worksheet (or each page if a sheet has multiple pages) becomes a frame in the GIF.
            workbook.Save(outputGif, SaveFormat.Gif);

            Console.WriteLine($"Animated GIF created successfully at '{outputGif}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
