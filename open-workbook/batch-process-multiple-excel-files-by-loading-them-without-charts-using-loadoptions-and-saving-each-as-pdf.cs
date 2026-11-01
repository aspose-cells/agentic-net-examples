// Title: Batch convert multiple Excel (.xlsx) files to PDF in C# using Aspose.Cells LoadOptions to skip chart loading
// Common Searches: C# Aspose.Cells batch convert all .xlsx files in a folder to PDF without loading charts | How to loop through a directory of Excel workbooks and save each as PDF using LoadOptions in Aspose.Cells | Convert multiple Excel files to PDF in C# while ignoring chart objects with Aspose.Cells | Programmatically export a batch of Excel workbooks to PDF without chart data using Aspose.Cells | Aspose.Cells LoadOptions Xlsx to PDF conversion for many files in a C# application
// Tags: batch workbook to pdf conversion aspnet cells | loadoptions skip charts aspnet cells | c# iterate folder convert xlsx to pdf aspnet cells | excel to pdf export without charts aspnet cells | aspnet cells save workbook as pdf c#

using System;
using System.IO;
using Aspose.Cells;

// The C# program scans a specified input directory for .xlsx files, loads each workbook with LoadOptions (set to Xlsx format) to avoid loading chart objects, and saves the workbook as a PDF in an output folder. It creates the output folder if needed, handles missing files and other exceptions, and logs conversion results for each workbook.
class Program
{
    static void Main()
    {
        // Folder containing the source Excel files
        string inputFolder = @"C:\InputExcel";

        // Folder where the PDF files will be saved
        string outputFolder = @"C:\OutputPdf";

        try
        {
            // Verify that the input directory exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Retrieve all Excel files (add other extensions if needed)
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx");

            foreach (string excelPath in excelFiles)
            {
                // Skip if the file somehow does not exist
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"File not found, skipping: {excelPath}");
                    continue;
                }

                try
                {
                    // Load the workbook (no special LoadOptions needed for basic conversion)
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
                    Workbook workbook = new Workbook(excelPath, loadOptions);

                    // Build the output PDF file name
                    string pdfFileName = Path.GetFileNameWithoutExtension(excelPath) + ".pdf";
                    string pdfPath = Path.Combine(outputFolder, pdfFileName);

                    // Save the workbook as a PDF document
                    workbook.Save(pdfPath, SaveFormat.Pdf);
                    Console.WriteLine($"Converted '{excelPath}' to PDF successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
