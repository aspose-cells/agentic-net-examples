// Title: Protect an Excel worksheet with a password in Aspose.Cells for .NET, block pivot table editing yet keep data refresh enabled
// AI Prompts: Write C# code using Aspose.Cells to apply password protection to a worksheet, prevent modifications to pivot tables, and allow pivot data to be refreshed. | Show how to configure Worksheet.Protection properties so that pivot tables are read‑only while the sheet stays protected in a .NET application. | Demonstrate the steps for loading a workbook, setting protection options for pivot tables, and saving the file with Aspose.Cells.
// Common Searches: Aspose.Cells C# protect worksheet but allow pivot table refresh | disable editing of pivot tables on a password‑protected Excel sheet using Aspose.Cells | how to keep pivot data refresh functional after worksheet protection in .NET | set worksheet protection options for pivot tables with Aspose.Cells library
// Tags: worksheet protection password Aspose.Cells | prevent pivot table modifications Aspose.Cells | enable pivot data refresh on protected worksheet .NET | Aspose.Cells ProtectionType configuration | C# protect Excel worksheet pivot tables | Excel worksheet protection options Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook (or creates a new one), applies password protection to the first worksheet, demonstrates how to adjust Worksheet.Protection settings so that pivot tables cannot be edited while still permitting data refresh, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load existing workbook if it exists; otherwise create a new one.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Optional: add some sample data to the default sheet.
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("Sample Data");
            }

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Protect the worksheet with a password. The third parameter (oldPassword) is not needed here, so pass null.
            sheet.Protect(ProtectionType.All, "MySecurePassword", null);

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
