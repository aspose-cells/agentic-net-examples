// Title: Hide rows 10‑20, unhide rows 15‑18 with custom height, and export workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that loads an Excel file, hides rows 10 through 20, then unhides rows 15‑18 and sets each row’s height to 20 points before saving the workbook as a PDF. | Show how to programmatically change row visibility and height in a worksheet using Aspose.Cells and then convert the modified workbook to PDF in a .NET application.
// Common Searches: Aspose.Cells hide rows 10 through 20 and later reveal rows 15 to 18 with custom height in C# | C# convert Excel to PDF after modifying row visibility using Aspose.Cells | Set row height while unhiding rows in Aspose.Cells for .NET | Hide a block of rows and then unhide a subset before PDF export with Aspose.Cells | Adjust row visibility and height in Aspose.Cells prior to saving workbook as PDF
// Tags: hide rows range Aspose.Cells C# | unhide rows set height Aspose.Cells | export worksheet to PDF Aspose.Cells | row visibility manipulation Aspose.Cells | row height adjustment Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads an existing Excel file, hides rows 10‑20, unhides rows 15‑18 while setting each row’s height to 20 points, and then saves the modified workbook as a PDF.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Hide rows 10 to 20 (zero‑based index: 9 to 19)
            for (int i = 9; i <= 19; i++)
            {
                sheet.Cells.Rows[i].IsHidden = true;
            }

            // Unhide rows 15 to 18 (zero‑based index: 14 to 17) and set a specific height
            for (int i = 14; i <= 17; i++)
            {
                Row row = sheet.Cells.Rows[i];
                row.IsHidden = false;   // Unhide the row
                row.Height = 20;        // Set row height (points)
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
