// Title: Remove write protection from a shared Excel workbook and save it unprotected with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# remove write protection from shared Excel workbook | How to unprotect an Excel file and save without password using Aspose.Cells .NET | C# code to load a protected workbook and save it unprotected with Aspose.Cells | Programmatically disable workbook protection in Excel using Aspose.Cells API | Aspose.Cells example for removing workbook password protection in .NET
// Tags: Aspose.Cells workbook unprotect method | C# remove Excel write protection | Aspose.Cells save workbook without password | shared Excel workbook unprotect .NET | Aspose.Cells handling workbook protection

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads a shared Excel workbook, calls Workbook.Unprotect with an empty password to clear any write protection, and then saves the workbook as a new file without any password protection.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "SharedWorkbook.xlsx";
            string outputPath = "SharedWorkbook_Unprotected.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Remove any write protection (empty password assumes none or removes it)
                workbook.Unprotect("");

                // Save the workbook without protection
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved without protection to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
