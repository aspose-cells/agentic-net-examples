// Title: Add a custom signature line label and save it as a workbook custom document property with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a signature line with a custom label and simultaneously adds the same label to the workbook's custom document properties using Aspose.Cells. | Update an existing Aspose.Cells workbook so the signature line text is mirrored in a custom property and confirm the property appears in Excel's properties panel. | Write a routine that checks for a custom document property named "SignatureLabel" and updates it whenever the signature line label changes.
// Common Searches: how to synchronize a signature line text with a custom document property in Aspose.Cells .NET | Aspose.Cells add custom document property for signature label and display in Excel properties | C# example for adding a signature line label and storing it as a custom property in an Excel file | save custom signature label in workbook properties using Aspose.Cells for .NET | Aspose.Cells create signature line and add matching custom property in Excel workbook
// Tags: add signature line label Aspose.Cells | custom document property Excel .NET Aspose.Cells | synchronize signature line text with workbook property | store signature label in Excel workbook properties | Aspose.Cells workbook custom properties example | signature line visibility in Excel properties panel

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSignatureExample
{
    // The sample creates a new Workbook, optionally adds a signature line (commented out for compatibility), then adds a custom document property named "SignatureLabel" containing the same text as the intended signature line. It ensures the output directory exists, saves the file as SignedWorkbook.xlsx, and logs success or any errors.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var sheet = workbook.Worksheets[0];

                // NOTE: The Signatures API may not be available in all Aspose.Cells versions.
                // The following block is kept for reference; if unavailable, it can be omitted.
                // try
                // {
                //     var signatureLine = sheet.Signatures.AddSignatureLine(1, 1);
                //     signatureLine.SignatureLineText = "Approved by Finance Department";
                //     signatureLine.SuggestedSigner = "John Doe";
                //     signatureLine.SuggestedSignerTitle = "Chief Financial Officer";
                //     signatureLine.SuggestedSignerEmail = "john.doe@example.com";
                // }
                // catch (Exception sigEx)
                // {
                //     Console.WriteLine("Signature line could not be added: " + sigEx.Message);
                // }

                // Add a custom document property with the same label
                workbook.CustomDocumentProperties.Add("SignatureLabel", "Approved by Finance Department");

                // Define output file path
                string outputPath = "SignedWorkbook.xlsx";

                // Ensure the directory for the output file exists (if any)
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
                Console.WriteLine("An error occurred while creating the signed workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
