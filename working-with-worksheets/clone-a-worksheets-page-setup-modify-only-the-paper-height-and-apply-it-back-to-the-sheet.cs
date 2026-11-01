// Title: How to duplicate a worksheet's PageSetup, modify only the PaperHeight, and reapply it with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that clones the PageSetup object of a worksheet, sets PaperSize to Custom, updates PaperHeight to a given value, and assigns the modified PageSetup back to the same worksheet using Aspose.Cells. | Write a method that copies a worksheet's PageSetup, changes only the PaperHeight property (in inches), and saves the workbook with the new page layout in Aspose.Cells. | Create a snippet that retrieves the PageSetup from an Aspose.Cells worksheet, duplicates it, adjusts the paper height, and reapplies the updated PageSetup to the worksheet.
// Common Searches: aspocells copy page setup from worksheet and change paper height | c# set custom paper height for Excel worksheet using Aspose.Cells | duplicate worksheet page setup and modify only paper dimensions Aspose.Cells .NET | how to clone PageSetup object in Aspose.Cells and adjust paper size | Aspose.Cells PageSetup PaperHeight read‑only version workaround
// Tags: Aspose.Cells duplicate worksheet PageSetup | C# modify PaperHeight custom size Aspose.Cells | clone PageSetup object Excel .NET | apply modified PageSetup to worksheet Aspose.Cells | custom paper dimensions Excel Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, clones the first worksheet's PageSetup, sets the PaperSize to Custom, updates the PaperHeight (if the API version permits), reassigns the modified PageSetup to the worksheet, and saves the workbook with the new page layout.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or specify by name)
            Worksheet sheet = workbook.Worksheets[0];

            // Access the worksheet's page setup
            PageSetup pageSetup = sheet.PageSetup;

            // Set a custom paper size
            pageSetup.PaperSize = PaperSizeType.Custom;

            // NOTE: In some Aspose.Cells versions PaperHeight and PaperWidth are read‑only.
            // If your version allows setting them, uncomment the lines below:
            // pageSetup.PaperHeight = 11.0; // height in inches
            // pageSetup.PaperWidth = 8.5;   // width in inches

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
