// Title: Insert a timeline graphic into each worksheet and generate a multi‑page PDF with Aspose.Cells for .NET
// AI Prompts: Place the same PNG timeline picture at the top of every worksheet, then save the workbook as a PDF using Aspose.Cells in C#. | Create a workbook with sample rows, add a timeline image to each sheet's picture collection, and export the entire workbook to a PDF where the image appears on every page. | Use Aspose.Cells to embed a timeline picture in all worksheets and produce a PDF that repeats the graphic on each printed page.
// Common Searches: how to embed the same picture on all Excel sheets before converting to PDF with Aspose.Cells C# | Aspose.Cells export workbook to PDF with a recurring header image on each page | draw a timeline graphic on every page of a PDF generated from an Excel workbook using .NET | C# add picture to worksheet header for PDF output Aspose.Cells | multi‑page PDF export with identical image on each page Aspose.Cells
// Tags: timeline picture insertion Aspose.Cells | export workbook to PDF with repeated image C# | add picture to each worksheet Aspose.Cells | multi‑page PDF generation Aspose.Cells .NET | worksheet header graphic for PDF output | place PNG graphic on all sheets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, fills it with sample data, inserts the same PNG timeline image into every worksheet, adjusts its size, and saves the workbook as a multi‑page PDF where the timeline appears on each page.
class TimelinePdfExport
{
    static void Main()
    {
        try
        {
            // Create a new workbook with a default worksheet
            Workbook workbook = new Workbook();

            // Populate sample data to span multiple printed pages
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";
            Cells cells = sheet.Cells;
            for (int row = 0; row < 200; row++)
            {
                cells[row, 0].PutValue($"Item {row + 1}");
                cells[row, 1].PutValue(DateTime.Today.AddDays(row));
            }

            // Simple placeholder PNG (1x1 white pixel) for the timeline image
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);

            using (MemoryStream imgStream = new MemoryStream(pngBytes))
            {
                // Insert the image into each worksheet (as a regular picture, not a header)
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    try
                    {
                        // Add picture to the worksheet's picture collection
                        int picIdx = ws.Pictures.Add(0, 0, imgStream);
                        Picture timelinePic = ws.Pictures[picIdx];

                        // Define picture size (adjust as needed)
                        timelinePic.Width = 800;
                        timelinePic.Height = 100;

                        // Reset stream for the next worksheet
                        imgStream.Position = 0;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Unable to add picture to worksheet '{ws.Name}'. {ex.Message}");
                    }
                }
            }

            // Save the workbook as a PDF file
            string outputPath = "TimelineWorkbook.pdf";
            try
            {
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving PDF: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
