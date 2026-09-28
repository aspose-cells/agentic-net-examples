// Title: Add a rectangle shape with compressed text to an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a new Workbook, inserts a rectangle shape on the first worksheet, assigns custom text to the shape, and saves the workbook with Aspose.Cells. | Explain how to emulate reduced character spacing for shape text in Aspose.Cells, including any work‑arounds or alternative font settings.
// Common Searches: how to set negative character spacing for shape text in Aspose.Cells C# | compress text inside a rectangle shape using Aspose.Cells .NET | Aspose.Cells shape font properties missing CharSpacing | add rectangle shape with custom text to Excel workbook using C# Aspose.Cells | workaround for lack of CharSpacing property in Aspose.Cells shapes
// Tags: aspose.cells insert rectangle shape c# | excel shape text spacing limitation aspose.cells | c# aspose.cells shape font configuration | generate excel workbook with shapes aspose.cells | aspose.cells shape text formatting options

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, accesses the first Worksheet, adds a rectangle shape of specified size, sets its text to "Compressed Text", notes that Aspose.Cells does not expose a CharSpacing property for shape fonts, and saves the file as CompressedTextShape.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle shape and obtain the created shape object
                Shape shape = sheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    2,    // upper left row (zero‑based)
                    2,    // upper left column (zero‑based)
                    0,    // top offset in points
                    0,    // left offset in points
                    60,   // height in points
                    300); // width in points

                // Set the text that will appear inside the shape
                shape.Text = "Compressed Text";

                // Note: Aspose.Cells does not provide a CharSpacing property for shape fonts.
                // Additional font settings can be applied here if needed, e.g.:
                // shape.Font.Size = 12;

                // Define output file path
                string outputPath = "CompressedTextShape.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
