// Title: Insert a Logo and Freeze Top Rows in Excel with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, inserts a PNG logo that spans cells A1:C3 (if the file exists), freezes the first three rows to keep the branding visible, and saves the file as LogoWithFreeze.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells insert image | Aspose.Cells freeze panes | C# add logo to Excel | picture spanning cells Aspose.Cells | freeze top rows Aspose.Cells | Excel branding C# | Workbook.Save Aspose.Cells | FreezePanes method C# | Aspose.Cells picture Add
// Common Searches: Aspose.Cells add logo to Excel sheet C# | How to freeze rows after inserting picture Aspose.Cells | Insert image spanning multiple cells Aspose.Cells .NET | Freeze top rows in workbook using Aspose.Cells | C# Aspose.Cells picture and FreezePanes example
// Developer Intent: Add a logo image to the worksheet header and keep it visible by freezing the top rows.
// Use Cases: Design a branded report template where the company logo stays fixed while scrolling through data. | Generate invoices with a top‑left logo and frozen header rows for easy navigation. | Build a dashboard workbook that displays a branding banner and locks it in place as users scroll.
// AI Prompts: Write C# code with Aspose.Cells that inserts a PNG logo spanning A1:C3 and then freezes the first three rows. | Provide an Aspose.Cells example that checks for a logo file, adds it to the worksheet, and applies FreezePanes to keep the branding visible. | Create a reusable C# method that accepts a workbook, image path, and row count, inserts the image at the top of the sheet, and freezes the specified rows.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace InsertLogoAndFreezeRowsApp
{
    // Creates a new workbook, inserts a PNG logo that spans cells A1:C3 (if the file exists), freezes the first three rows to keep the branding visible, and saves the file as LogoWithFreeze.xlsx using Aspose.Cells for .NET.
    class InsertLogoAndFreezeRows
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                string logoPath = "logo.png";

                // Insert logo picture if the file exists
                if (File.Exists(logoPath))
                {
                    // Add picture spanning rows 0‑2 and columns 0‑2 (cells A1:C3)
                    worksheet.Pictures.Add(0, 0, 2, 2, logoPath);
                }
                else
                {
                    Console.WriteLine($"Warning: Logo file '{logoPath}' not found. Skipping picture insertion.");
                }

                // Freeze the first three rows so the branding stays visible
                worksheet.FreezePanes(3, 0, 3, 0);

                // Save the workbook
                string outputPath = "LogoWithFreeze.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
