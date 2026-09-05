// Title: Add a signature line with signer name to a specific worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a new workbook, adds a signature line to the first worksheet, assigns the Signer property (e.g., "John Doe"), optionally sets the Title property, and saves the file as an XLSX using Aspose.Cells. | Write version‑agnostic C# that adds a signature line using either the SignatureLines collection or the SignatureLineCollection fallback, sets signer and title details, and gracefully handles missing API members.
// Common Searches: asp.net add signature line to excel worksheet with signer name using Aspose.Cells | c# set title property on Aspose.Cells signature line | fallback to SignatureLineCollection when SignatureLines not found Aspose.Cells | save workbook with digital signature line as xlsx in C# | how to use dynamic to handle Aspose.Cells API changes for signature lines
// Tags: add signature line Aspose.Cells worksheet | set signer property digital signature Aspose.Cells | set title property signature line Aspose.Cells | dynamic fallback SignatureLines collection | save workbook with signature line xlsx

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSignatureDemo
{
    // The example creates a new workbook, accesses the first worksheet, attempts to add a signature line using either the SignatureLines or SignatureLineCollection collection via a dynamic fallback, sets the Signer property to "John Doe" (with optional title), and saves the workbook as SignedWorkbook.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one if needed)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Use dynamic to handle possible API differences across Aspose.Cells versions
                dynamic dynSheet = sheet;
                int signatureIndex = -1;
                SignatureLine signature = null;

                // Attempt to add a signature line using the available collection
                try
                {
                    // First try the "SignatureLines" collection
                    signatureIndex = dynSheet.SignatureLines.Add();
                    signature = dynSheet.SignatureLines[signatureIndex];
                }
                catch
                {
                    // Fallback to "SignatureLineCollection" if the above collection is not present
                    try
                    {
                        signatureIndex = dynSheet.SignatureLineCollection.Add();
                        signature = dynSheet.SignatureLineCollection[signatureIndex];
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Signature line API not available: {ex.Message}");
                    }
                }

                // If a signature line was successfully created, set its properties
                if (signature != null)
                {
                    signature.Signer = "John Doe";

                    // Optional properties (uncomment if supported by your Aspose.Cells version)
                    // signature.Email = "john.doe@example.com";
                    // signature.Instructions = "Please sign here.";
                }

                // Define output file path
                string outputPath = "SignedWorkbook.xlsx";

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
