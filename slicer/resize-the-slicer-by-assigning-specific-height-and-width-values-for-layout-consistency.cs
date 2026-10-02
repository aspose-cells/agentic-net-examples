// Title: Resize an Excel worksheet slicer by setting explicit width and height with Aspose.Cells for .NET (C#)
// AI Prompts: Set the first worksheet slicer's Width to 200 points and Height to 100 points using Aspose.Cells in C#. | Verify that a worksheet contains slicers before applying size changes with Aspose.Cells. | Save the workbook after adjusting slicer dimensions to keep the layout consistent.
// Common Searches: how to change slicer size in Excel using Aspose.Cells C# | Aspose.Cells set slicer width and height programmatically | resize Excel slicer dimensions .NET workbook example
// Tags: Aspose.Cells slicer dimension setting | C# adjust slicer layout | Excel slicer size modification Aspose | worksheet slicer resizing .NET | programmatic slicer size control

using Aspose.Cells;
using Aspose.Cells.Slicers;
using System;
using System.IO;

// The code loads an existing workbook, checks for slicers on the first worksheet, sets the first slicer's Width to 200 points and Height to 100 points, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the slicer collection from the worksheet
            SlicerCollection slicers = sheet.Slicers;

            // Resize the first slicer if any are present
            if (slicers.Count > 0)
            {
                Slicer slicer = slicers[0];
                slicer.Width = 200;   // Desired width in points
                slicer.Height = 100;  // Desired height in points
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
