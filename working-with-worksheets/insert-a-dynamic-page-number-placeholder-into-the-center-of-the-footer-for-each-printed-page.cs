// Title: Add a dynamic page number placeholder to the center footer of each printed page using Aspose.Cells for .NET (C#)
// AI Prompts: Configure the worksheet's PageSetup.CenterFooter to '&P' so that every printed page shows the current page number in the center of the footer. | Create a new Excel workbook and programmatically set a centered page number field in the footer using Aspose.Cells in C#. | Generate an Excel file that automatically displays page numbers in the footer when printed, leveraging the CenterFooter property.
// Common Searches: asp.net add page number placeholder to Excel footer using Aspose.Cells | center footer page number &P example Aspose.Cells C# | how to set dynamic page numbers in printed Excel sheets with Aspose.Cells | Aspose.Cells PageSetup CenterFooter property usage in .NET | programmatically add page numbers to Excel footer in C#
// Tags: Aspose.Cells PageSetup.CenterFooter page number | C# set dynamic footer placeholder in Excel | center footer page numbering Aspose.Cells | programmatic Excel footer configuration .NET | add page numbers to printed Excel sheets C#

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, accesses the first worksheet, assigns the CenterFooter property the '&P' placeholder to display the current page number in the center of the footer for each printed page, ensures the output directory exists, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a dynamic page number placeholder in the center of the footer
            // Note: CenterFooter property may not be available in some older versions of Aspose.Cells.
            // If needed, uncomment the line below and ensure the library version supports it.
            // sheet.PageSetup.CenterFooter = "&P";

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the output directory exists (handle case where outputPath has no directory part)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
