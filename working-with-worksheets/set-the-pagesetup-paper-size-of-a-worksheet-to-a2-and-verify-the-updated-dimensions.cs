// Title: How to set a worksheet’s page‑setup paper size to A2 and confirm it using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set worksheet page setup to A2 paper size | assign A2 paper size with numeric enum value in Aspose.Cells | verify worksheet page setup paper size after setting it in .NET | save workbook after changing page setup paper size with Aspose.Cells | create in‑memory workbook and apply A2 paper size using Aspose.Cells C#
// Tags: worksheet page setup paper size Aspose.Cells | cast numeric value to PaperSizeType C# | validate page setup property after assignment Aspose.Cells | persist page setup changes by saving workbook Aspose.Cells | in‑memory workbook creation Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The example creates an in‑memory workbook, accesses the first worksheet, casts the numeric value 2 to PaperSizeType to represent A2, assigns it to the worksheet's PageSetup.PaperSize, checks that the assignment succeeded, saves the workbook as "PageSetup_A2.xlsx" to persist the setting, and outputs status messages while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (in‑memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // PaperSizeType does not expose A2 in some older versions.
            // Use the numeric value (2) that corresponds to A2.
            const int a2NumericValue = 2;
            PaperSizeType a2Size = (PaperSizeType)a2NumericValue;

            // Set the page‑setup paper size to A2
            sheet.PageSetup.PaperSize = a2Size;

            // Verify that the paper size was updated
            if (sheet.PageSetup.PaperSize == a2Size)
            {
                Console.WriteLine("Paper size successfully set to A2.");
            }
            else
            {
                Console.WriteLine("Failed to set paper size to A2.");
            }

            // Optional: save the workbook to verify the setting persists
            string outputPath = "PageSetup_A2.xlsx";

            // Ensure the directory exists before saving
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
