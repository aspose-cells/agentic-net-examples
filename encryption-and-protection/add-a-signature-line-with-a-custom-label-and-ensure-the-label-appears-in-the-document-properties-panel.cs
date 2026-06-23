using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    class AddSignatureLineWithLabel
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Define a signature line with its properties
                SignatureLine signatureLine = new SignatureLine
                {
                    Signer = "John Doe",
                    Title = "Project Manager",
                    Email = "john.doe@example.com",
                    Instructions = "Please sign this document",
                    IsLine = true,
                    AllowComments = true,
                    ShowSignedDate = true
                };

                // Add the signature line to cell B2 (row 1, column 1 – zero‑based)
                worksheet.Shapes.AddSignatureLine(1, 1, signatureLine);

                // Add a custom document property to store the label (string type)
                workbook.CustomDocumentProperties.Add("SignatureLabel", signatureLine.Instructions);

                // Save the workbook
                string outputPath = "SignatureLineWithCustomLabel.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}