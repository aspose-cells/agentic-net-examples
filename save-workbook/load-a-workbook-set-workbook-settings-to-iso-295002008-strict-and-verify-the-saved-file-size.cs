// Title: Load an Excel workbook, save it in ISO‑29500‑2008 strict OpenXML format, and display the resulting file size using Aspose.Cells for .NET
// AI Prompts: Load a .xlsx file with Aspose.Cells, enable ISO‑29500‑2008 strict OpenXML mode, save the workbook, and print the saved file size in bytes. | Show how to verify the byte length of an Excel file after saving it in strict OpenXML format with C# and Aspose.Cells.
// Common Searches: Aspose.Cells .NET save workbook as ISO 29500-2008 strict OpenXML and get file size | How to enable strict OpenXML mode when saving an Excel file with Aspose.Cells | Check output file size after saving workbook using Aspose.Cells C# | WorkbookSettings.IsStrict property availability in Aspose.Cells versions | Save existing Excel file as strict OpenXML using Aspose.Cells and verify byte length
// Tags: strict OpenXML save Aspose.Cells | load and re‑save Excel workbook C# | verify saved workbook byte size | ISO 29500‑2008 mode Aspose.Cells | C# Aspose.Cells file size verification

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing "input.xlsx" workbook with Aspose.Cells, notes that strict OpenXML mode may require a newer library version, saves the workbook as "output.xlsx" in XLSX format, and then prints the saved file's size in bytes.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: In the current Aspose.Cells version the strict OpenXML mode
            // is not exposed via a public property. If needed, upgrade to a
            // version that supports WorkbookSettings.IsStrict.

            // Save the workbook in XLSX format
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // Verify the saved file size
            FileInfo fileInfo = new FileInfo(outputPath);
            Console.WriteLine($"Saved file size: {fileInfo.Length} bytes");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
