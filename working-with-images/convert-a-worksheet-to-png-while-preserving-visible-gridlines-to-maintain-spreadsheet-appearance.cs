// Title: Export Excel Worksheet to PNG with Visible Gridlines using Aspose.Cells for .NET
// Description: Loads a workbook (or creates one), turns on worksheet gridlines, sets ImageOrPrintOptions with a dotted gridline style, and renders the first page to a PNG file named "WorksheetOutput.png".
// Keywords: Aspose.Cells PNG export | Excel worksheet image with gridlines | C# SheetRender ToImage example | ImageOrPrintOptions GridlineType | convert Excel to PNG .NET
// Common Searches: Aspose.Cells export worksheet to PNG with gridlines | C# render Excel sheet as PNG preserving gridlines | how to keep gridlines when converting Excel to image | SheetRender ToImage gridline settings | ImageOrPrintOptions gridline type example
// Developer Intent: Create a PNG image of a worksheet that includes the visible gridlines.
// Use Cases: Generate printable snapshots of reports while retaining the familiar grid layout. | Provide web‑based previews of spreadsheets with exact Excel appearance. | Automate batch conversion of multiple sheets to PNG for documentation or UI thumbnails.
// AI Prompts: Write C# code with Aspose.Cells to export a specific worksheet to PNG and keep gridlines visible. | Show how to configure ImageOrPrintOptions to use a dotted gridline style when rendering a sheet to PNG. | Explain how to change the gridline type or output format while preserving the worksheet's visual fidelity.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Loads a workbook (or creates one), turns on worksheet gridlines, sets ImageOrPrintOptions with a dotted gridline style, and renders the first page to a PNG file named "WorksheetOutput.png".
    public class WorksheetToPngWithGridlines
    {
        public static void Run()
        {
            try
            {
                const string inputPath = "InputWorkbook.xlsx";

                // Load existing workbook if it exists; otherwise create a new workbook
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Ensure the workbook is disposed properly
                using (workbook)
                {
                    // Access the first worksheet
                    Worksheet worksheet = workbook.Worksheets[0];

                    // Make gridlines visible
                    worksheet.IsGridlinesVisible = true;

                    // Configure image rendering options
                    ImageOrPrintOptions options = new ImageOrPrintOptions
                    {
                        // Default format is PNG; explicit setting omitted to avoid API mismatch
                        GridlineType = GridlineType.Dotted
                    };

                    // Render the first page of the worksheet to a PNG file
                    SheetRender sheetRender = new SheetRender(worksheet, options);
                    sheetRender.ToImage(0, "WorksheetOutput.png");

                    Console.WriteLine("Worksheet rendered to PNG with gridlines successfully.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            WorksheetToPngWithGridlines.Run();
        }
    }
}
