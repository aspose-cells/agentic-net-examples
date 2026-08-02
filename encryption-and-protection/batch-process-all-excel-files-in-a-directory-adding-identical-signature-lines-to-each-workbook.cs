// Title: Batch add identical signature lines to Excel workbooks with Aspose.Cells (C#)
// Description: Scans a folder for .xlsx, .xls, .xlsm, .xlsb files, loads each workbook with Aspose.Cells, inserts the same SignatureLine into cell A1 of the first worksheet, and saves the signed copy to an output directory while handling missing files and runtime errors.
// Keywords: Aspose.Cells | C# signature line | batch add signature to Excel | process multiple workbooks | add signature line programmatically | Excel folder processing | digital signature Excel C# | signature line Aspose.Cells | automate Excel signing
// Common Searches: add signature line to all Excel files in a folder C# | batch insert signature line Aspose.Cells | how to automate signing of multiple workbooks | C# script to add identical signature line to Excel workbooks | Aspose.Cells example for batch signature
// Developer Intent: Insert the same signature line into each workbook in a directory and save the signed copies.
// Use Cases: Sign a set of monthly financial statements before distribution | Apply a compliance approval stamp to a batch of invoices | Embed an approval line in all template files for data collection | Prepare legally signed reports for audit by adding a signature line to each file
// AI Prompts: Write C# code using Aspose.Cells to add a configurable signature line to every worksheet in every workbook within a specified folder. | Recommend robust logging and exception handling strategies for a batch Excel signing utility. | Show how to customize the signer name, title, or email based on each file’s name while adding signature lines. | Explain how to add a visible signature image together with a signature line in a batch process.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Scans a folder for .xlsx, .xls, .xlsm, .xlsb files, loads each workbook with Aspose.Cells, inserts the same SignatureLine into cell A1 of the first worksheet, and saves the signed copy to an output directory while handling missing files and runtime errors.
class BatchSignatureAdder
{
    static void Main()
    {
        try
        {
            // Input and output directories
            string inputDir = @"C:\InputExcelFiles";
            string outputDir = @"C:\SignedExcelFiles";

            // Verify input directory exists
            if (!Directory.Exists(inputDir))
            {
                Console.WriteLine($"Input directory not found: {inputDir}");
                return;
            }

            // Ensure the output directory exists
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Supported Excel extensions
            string[] extensions = new[] { "*.xlsx", "*.xls", "*.xlsm", "*.xlsb" };

            // Process each file matching the extensions
            foreach (string ext in extensions)
            {
                foreach (string filePath in Directory.GetFiles(inputDir, ext))
                {
                    try
                    {
                        // Verify the file exists before loading
                        if (!File.Exists(filePath))
                        {
                            Console.WriteLine($"File not found: {filePath}");
                            continue;
                        }

                        // Load the workbook
                        Workbook workbook = new Workbook(filePath);

                        // Use the first worksheet; adjust if needed
                        Worksheet worksheet = workbook.Worksheets[0];

                        // Configure a signature line (identical for all workbooks)
                        SignatureLine signatureLine = new SignatureLine
                        {
                            Signer = "John Doe",
                            Title = "Approved",
                            Email = "john.doe@example.com",
                            Instructions = "Please sign to confirm the content.",
                            AllowComments = true,
                            ShowSignedDate = true,
                            IsLine = true
                        };

                        // Add the signature line at cell A1 (row 0, column 0)
                        worksheet.Shapes.AddSignatureLine(0, 0, signatureLine);

                        // Save the modified workbook to the output directory
                        string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
                        workbook.Save(outputPath);
                        Console.WriteLine($"Signed file saved: {outputPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
