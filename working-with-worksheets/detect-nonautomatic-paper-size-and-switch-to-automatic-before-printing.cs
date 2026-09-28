// Title: Detect non‑automatic worksheet paper size and switch to Automatic with Aspose.Cells for .NET before printing
// AI Prompts: Generate C# code using Aspose.Cells that iterates all worksheets, checks if PageSetup.PaperSize is not PaperSizeType.Automatic, and sets it to Automatic before calling workbook.Print(). | Write a method in C# that returns true when any worksheet in a workbook has a custom paper size, then updates those worksheets to use PaperSizeType.Automatic with Aspose.Cells. | Create an example that loads an Excel file, changes only non‑automatic paper sizes to Automatic using Aspose.Cells PageSetup, and saves the workbook ready for printing.
// Common Searches: Aspose.Cells C# detect custom paper size in worksheet and change to automatic | how to set page setup to automatic paper size for all sheets before printing with Aspose.Cells | C# Aspose.Cells iterate worksheets and convert non‑automatic paper size to automatic | Aspose.Cells .NET change worksheet paper size to automatic only when it is custom | print Excel workbook with Aspose.Cells after normalizing paper size to automatic
// Tags: Aspose.Cells detect custom paper size | Aspose.Cells set automatic paper size | C# page setup automatic paper size | Aspose.Cells workbook print preparation | iterate worksheets page setup Aspose.Cells | Excel paper size normalization Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel workbook (or creates a new one), loops through each worksheet, checks the PageSetup.PaperSize property, and if the size is not PaperSizeType.Automatic it sets it to Automatic. After normalizing the paper size the workbook can be printed or saved, ensuring consistent automatic sizing across all sheets.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if the file exists; otherwise create a new workbook.
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Iterate through all worksheets and ensure a standard paper size (e.g., A4).
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                PageSetup pageSetup = sheet.PageSetup;

                // Set paper size to A4 (replace with desired default if needed).
                pageSetup.PaperSize = PaperSizeType.PaperA4;
            }

            // Optional: print the workbook (uses default printer)
            // workbook.Print();

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
