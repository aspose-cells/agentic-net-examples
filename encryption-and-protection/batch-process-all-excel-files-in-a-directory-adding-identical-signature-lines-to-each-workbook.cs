using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace BatchSignature
{
    class Program
    {
        static void Main(string[] args)
        {
            // Source directory containing Excel files
            string sourceDir = @"C:\ExcelFiles\Input";
            // Destination directory for signed files
            string destDir = @"C:\ExcelFiles\Signed";

            try
            {
                // Verify source directory exists
                if (!Directory.Exists(sourceDir))
                {
                    Console.WriteLine($"Source directory does not exist: {sourceDir}");
                    return;
                }

                // Ensure destination directory exists
                Directory.CreateDirectory(destDir);

                // Get all .xlsx, .xls, .xlsm files in the source directory
                string[] excelFiles = Directory.GetFiles(sourceDir, "*.*", SearchOption.TopDirectoryOnly);
                foreach (string filePath in excelFiles)
                {
                    // Process only supported Excel formats
                    string ext = Path.GetExtension(filePath).ToLowerInvariant();
                    if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm")
                        continue;

                    // Verify the file still exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook
                        using (Workbook workbook = new Workbook(filePath))
                        {
                            // Add the signature line to the first worksheet (customize as needed)
                            AddSignatureLine(workbook.Worksheets[0]);

                            // Build output file path (overwrite original name in destination folder)
                            string outputPath = Path.Combine(destDir, Path.GetFileName(filePath));

                            // Save the signed workbook
                            workbook.Save(outputPath);
                        }

                        Console.WriteLine($"Signed file saved: {Path.GetFileName(filePath)}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Adds an identical signature line to the specified worksheet
        private static void AddSignatureLine(Worksheet worksheet)
        {
            // Create a SignatureLine object and set its properties
            SignatureLine signatureLine = new SignatureLine
            {
                Signer = "John Doe",
                Title = "Approver",
                Email = "john.doe@example.com",
                Instructions = "Please sign to confirm the content.",
                AllowComments = true,
                ShowSignedDate = true,
                IsLine = true
            };

            // Define the position (top row and left column) where the signature line will be placed
            int topRow = 20;      // Row index (0‑based)
            int leftColumn = 2;   // Column index (0‑based)

            // Add the signature line to the worksheet's shape collection
            worksheet.Shapes.AddSignatureLine(topRow, leftColumn, signatureLine);
        }
    }
}