// Title: How to set a worksheet's right margin to 10 points using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set worksheet right margin to 10 points | C# page setup right margin property in Aspose.Cells example | Increase right margin for Excel sheet using Aspose.Cells .NET | Configure Excel page margins for long sentences with Aspose.Cells | Set RightMargin property in Aspose.Cells workbook programmatically
// Tags: Aspose.Cells right margin points | C# page setup right margin Excel | adjust worksheet right margin Aspose.Cells | Excel page margins configuration .NET | RightMargin property Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Creates a new workbook, accesses the first worksheet, sets the page setup RightMargin to 10 points to provide space for long sentences, ensures the output directory exists, and saves the file as Output.xlsx.
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

            // Set the right margin to 10 points (provides space for long sentences)
            // In Aspose.Cells the property is RightMargin (measured in points)
            sheet.PageSetup.RightMargin = 10;

            // Define output file path
            string outputPath = "Output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
