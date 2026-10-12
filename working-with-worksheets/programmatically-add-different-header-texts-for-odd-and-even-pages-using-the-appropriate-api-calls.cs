// Title: How to set different odd and even page headers in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells Workbook and Worksheet objects to assign a custom odd‑page header and a separate even‑page header with specific font settings. | Demonstrate how to call PageSetup.SetHeader with the correct index values to configure odd and even headers in an Excel file and then save the workbook. | Provide a complete .NET example that creates a new workbook, sets formatted odd/even headers, ensures the output folder exists, and saves the file as HeadersOddEven.xlsx.
// Common Searches: Aspose.Cells C# set odd page header and even page header in the same worksheet | PageSetup.SetHeader method index 0 odd 1 even example | How to apply font name and size to odd and even headers using Aspose.Cells | Create Excel file with different headers for odd and even pages in .NET
// Tags: Aspose.Cells PageSetup.SetHeader odd even | C# set odd page header Aspose.Cells | C# set even page header Aspose.Cells | Excel custom odd/even headers Aspose.Cells | Aspose.Cells save workbook with custom headers

using Aspose.Cells;
using System;
using System.IO;

// Creates a new Workbook, accesses the first Worksheet, and uses PageSetup.SetHeader(0, ...) and SetHeader(1, ...) to assign formatted odd and even page headers before saving the file as HeadersOddEven.xlsx.
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

            // Set odd and even page headers using SetHeader (0 = odd, 1 = even)
            // Header format: &"FontName,Style"&Size Text
            sheet.PageSetup.SetHeader(0, "&\"Arial,Bold\"&12 Odd Page Header");
            sheet.PageSetup.SetHeader(1, "&\"Arial,Bold\"&12 Even Page Header");

            // Define output file path
            string outputPath = "HeadersOddEven.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
