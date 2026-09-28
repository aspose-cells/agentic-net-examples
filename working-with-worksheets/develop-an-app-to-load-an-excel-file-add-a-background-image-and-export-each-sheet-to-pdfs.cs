// Title: Add a PNG background image to each worksheet and export worksheets as separate PDFs using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xlsx workbook with Aspose.Cells, inserts a PNG image as a free‑floating picture on every worksheet, and saves each worksheet individually as a PDF file in a target folder. | Show how to loop through all worksheets in a Workbook, apply a background picture, set the active sheet, and call Workbook.Save with SaveFormat.Pdf for each sheet using Aspose.Cells.
// Common Searches: aspnet add same background image to all Excel worksheets and export each sheet to PDF | c# aspose.cells insert picture as worksheet background then save each sheet as separate pdf | how to batch convert Excel worksheets to PDFs with a watermark image using Aspose.Cells | asp.net core load workbook, add png to each sheet, export each sheet to pdf
// Tags: insert png as worksheet background Aspose.Cells | worksheet-to-pdf conversion Aspose.Cells | freefloating picture placement Aspose.Cells | iterate worksheets set active sheet C# | batch excel sheet PDF export with image

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# program that loads an Excel file, adds a PNG image as a free‑floating background picture to every worksheet, and saves each worksheet as an individual PDF in a specified output folder using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Paths (adjust as needed)
        string excelPath = "input.xlsx";          // Excel file to load
        string backgroundImagePath = "bg.png";    // Background image file
        string outputFolder = "output";           // Folder for PDF files

        // Ensure output directory exists
        Directory.CreateDirectory(outputFolder);

        // Verify required files exist
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Error: Excel file not found at '{excelPath}'.");
            return;
        }

        if (!File.Exists(backgroundImagePath))
        {
            Console.WriteLine($"Error: Background image not found at '{backgroundImagePath}'.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Add background image as a picture covering the sheet
                    // The Add method returns the picture index; retrieve the Picture object.
                    int picIndex = sheet.Pictures.Add(0, 0, 0, 0, backgroundImagePath);
                    Picture pic = sheet.Pictures[picIndex];
                    pic.Placement = PlacementType.FreeFloating;
                }
                catch (Exception imgEx)
                {
                    Console.WriteLine($"Warning: Could not add background image to sheet '{sheet.Name}': {imgEx.Message}");
                }

                // Set the current sheet as the active sheet for saving
                workbook.Worksheets.ActiveSheetIndex = sheet.Index;

                // Build PDF file name based on sheet name
                string pdfFileName = $"{sheet.Name}.pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                try
                {
                    // Export the active sheet to PDF
                    workbook.Save(pdfPath, SaveFormat.Pdf);
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Error saving PDF for sheet '{sheet.Name}': {saveEx.Message}");
                }
            }

            Console.WriteLine("All sheets have been exported to PDF with background images.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
