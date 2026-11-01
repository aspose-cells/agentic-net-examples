// Title: Add a PNG picture to a specific cell in an Aspose.Cells workbook and export it as a single‑page TIFF using C#
// Common Searches: C# Aspose.Cells insert PNG image at cell Z100 and save workbook as TIFF | how to embed an overlay picture into an Excel worksheet with Aspose.Cells before TIFF conversion | export a single-page TIFF from a workbook that contains a picture using Aspose.Cells in C# | Aspose.Cells picture placement on worksheet then convert to TIFF in .NET | C# generate TIFF from Excel with embedded PNG overlay using Aspose.Cells
// Tags: add picture to worksheet cell Aspose.Cells | convert workbook with embedded image to TIFF C# | place PNG overlay in Excel sheet using Aspose.Cells | single-page TIFF export from Aspose.Cells workbook | C# image merging workflow with Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, writes text to cells A1 and A2, checks for an overlay PNG file, inserts the picture at cell Z100, and saves the worksheet as a single‑page TIFF named MergedResult.tiff, handling errors and disposing resources properly.
class WorkbookImageMerger
{
    static void Main()
    {
        Workbook workbook = null;
        try
        {
            // 1. Create a workbook and add some data
            workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells TIFF Example");
            sheet.Cells["A2"].PutValue(DateTime.Now.ToString());

            // 2. Verify the overlay image exists
            const string overlayPath = "overlay.png";
            if (!File.Exists(overlayPath))
            {
                Console.WriteLine($"Overlay image not found: {overlayPath}");
                return;
            }

            // 3. Insert the overlay picture into the worksheet (positioned at cell Z100)
            // Row and column indices are zero‑based
            sheet.Pictures.Add(99, 25, overlayPath);

            // 4. Save the workbook as a single‑page TIFF image
            const string resultPath = "MergedResult.tiff";
            workbook.Save(resultPath, SaveFormat.Tiff);
            Console.WriteLine($"Merged image saved to {resultPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        finally
        {
            // Cleanup Aspose.Cells objects
            workbook?.Dispose();
        }
    }
}
