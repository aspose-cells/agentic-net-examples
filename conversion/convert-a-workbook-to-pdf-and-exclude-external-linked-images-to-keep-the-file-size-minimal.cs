// Title: Convert an Excel workbook to PDF without embedded pictures using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove all worksheet pictures before saving as PDF | How to export Excel to PDF while excluding images with Aspose.Cells .NET | Minimize PDF size from Excel by stripping linked images using Aspose.Cells | Programmatically delete pictures from a workbook in C# prior to PDF conversion | Aspose.Cells save workbook as PDF without external image resources
// Tags: remove worksheet pictures Aspose.Cells C# | export workbook to PDF without images Aspose.Cells | minimize PDF size Excel Aspose.Cells | strip linked images during Excel to PDF conversion .NET | Aspose.Cells picture removal before PDF export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel file with Aspose.Cells, iterates through each worksheet to delete all pictures from the Pictures collection, and then saves the workbook as a PDF, resulting in a smaller PDF that contains no embedded or linked images.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string sourcePath = @"C:\Input\WorkbookWithLinkedImages.xlsx";

            // Path for the resulting PDF file
            string pdfPath = @"C:\Output\Workbook_Minimized.pdf";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(pdfPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Remove all pictures from each worksheet (Aspose.Cells does not expose IsLinked property)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate backwards to safely remove items from the collection
                for (int i = sheet.Pictures.Count - 1; i >= 0; i--)
                {
                    sheet.Pictures.RemoveAt(i);
                }
            }

            // Save the workbook as PDF
            workbook.Save(pdfPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to: {pdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
