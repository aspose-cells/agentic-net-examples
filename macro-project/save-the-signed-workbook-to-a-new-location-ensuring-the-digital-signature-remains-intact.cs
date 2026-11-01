// Title: Copy a digitally signed Excel workbook to a new location while preserving the signature with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save workbook without removing existing digital signature | How to duplicate a signed .xlsx file using Aspose.Cells .NET API | Preserve Excel digital signature when copying file programmatically in C# | Copy signed workbook to another path with Aspose.Cells while keeping signature intact | Save signed Excel workbook to new file location using Aspose.Cells SaveFormat.Xlsx
// Tags: Aspose.Cells preserve digital signature on save | C# copy signed Excel workbook | Workbook.Save preserving signature | SaveFormat.Xlsx digital signature retention | Aspose.Cells workbook cloning without signature loss

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookSignatureExample
{
    // The example loads an existing .xlsx file that may contain a digital signature, then saves it to a different path using Workbook.Save with SaveFormat.Xlsx, ensuring the original signature remains intact. If the source file is missing, a placeholder workbook is created first.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = @"C:\Path\SignedWorkbook.xlsx";
            string outputPath = @"C:\Path\SignedWorkbook_Copy.xlsx";

            try
            {
                // Ensure the input file exists; create a placeholder if it does not.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    Workbook placeholder = new Workbook();
                    placeholder.Save(inputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Created placeholder workbook at {inputPath}");
                }

                // Load the workbook (which may already contain a digital signature)
                Workbook workbook = new Workbook(inputPath);

                // Save without modifications to preserve any existing digital signature
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook successfully saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
