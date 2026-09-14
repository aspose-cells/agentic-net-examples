// Title: Copy a worksheet's PageSetup, adjust only the paper height, and reassign it using Aspose.Cells for .NET
// AI Prompts: Write C# code that clones the PageSetup object of a worksheet, changes the paper height to match Letter dimensions, and assigns the modified setup back to the same sheet with Aspose.Cells. | Show how to duplicate all page‑setup settings from one worksheet, modify only the paper height, and apply the updated PageSetup to another worksheet in a .NET workbook.
// Common Searches: asp.net aspose.cells copy page setup from one sheet to another and change paper size | how to change only the paper height of an Excel worksheet using Aspose.Cells C# | duplicate worksheet page setup settings without losing other properties Aspose.Cells | set paper size to Letter for a specific worksheet while preserving existing page setup Aspose.Cells
// Tags: clone worksheet page setup Aspose.Cells | modify paper height without affecting other page setup properties | apply copied page setup to another worksheet C# | configure worksheet for Letter paper dimensions Aspose.Cells | duplicate page setup settings between worksheets

using System;
using System.IO;
using Aspose.Cells;

// The example loads (or creates) a workbook, clones the first worksheet's PageSetup, changes only the paper height to Letter size, reassigns the modified PageSetup to the sheet, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load the workbook if the input file exists; otherwise create a new workbook.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // default workbook with one worksheet
            }

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Set the paper size to Letter (height = 11 inches). This changes the paper height.
            sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;

            // Save the modified workbook.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
