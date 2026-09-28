// Title: Add the same PNG watermark picture to every worksheet in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Insert a PNG image as a watermark on all worksheets of a workbook and save the file with Aspose.Cells in C#. | Load an existing Excel file, verify the watermark image exists, and programmatically add the picture to each sheet using Aspose.Cells. | Create a new workbook when the source file is missing, then apply the same picture shape to every worksheet while handling file‑not‑found errors in C#.
// Common Searches: C# Aspose.Cells how to add the same picture as a watermark to every sheet in an existing workbook | Aspose.Cells add PNG watermark to all worksheets programmatically | Load Excel file and apply picture shape to each worksheet using Aspose.Cells .NET | Handle missing watermark image when inserting picture into multiple worksheets with Aspose.Cells | Create new workbook if source file not found and then add watermark image to all sheets in C#
// Tags: add png watermark to all worksheets Aspose.Cells | insert picture shape across multiple Excel sheets C# | load workbook and apply image watermark Aspose.Cells | handle missing watermark file C# Aspose.Cells | fallback workbook creation when source file absent Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook (or creates a new one if the file is missing), checks that a PNG watermark image is present, adds that image as a picture to every worksheet, and saves the updated workbook while gracefully handling file‑not‑found and other exceptions.
class Program
{
    static void Main()
    {
        // Paths for input, output and watermark image
        string sourceWorkbookPath = "input.xlsx";
        string outputWorkbookPath = "output.xlsx";
        string watermarkImagePath = "watermark.png";

        try
        {
            // Verify that the watermark image exists
            if (!File.Exists(watermarkImagePath))
                throw new FileNotFoundException("Watermark image not found.", watermarkImagePath);

            Workbook workbook;

            // Load existing workbook if present; otherwise create a new one
            if (File.Exists(sourceWorkbookPath))
            {
                workbook = new Workbook(sourceWorkbookPath);
            }
            else
            {
                workbook = new Workbook(); // default workbook with one worksheet
            }

            // Apply the watermark image to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Add the image as a picture shape positioned at the top‑left corner
                    int pictureIndex = sheet.Pictures.Add(0, 0, watermarkImagePath);
                    // Optionally, you can adjust size/position here if needed
                    // Example: sheet.Pictures[pictureIndex].Width = 500;
                    // Example: sheet.Pictures[pictureIndex].Height = 300;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to add watermark on sheet '{sheet.Name}': {ex.Message}");
                }
            }

            // Save the modified workbook
            workbook.Save(outputWorkbookPath);
            Console.WriteLine($"Workbook saved successfully to '{outputWorkbookPath}'.");
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"File error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
