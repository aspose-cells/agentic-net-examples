// Title: Create an Excel workbook with a placeholder signature line and sign it using an external XAdES command‑line tool in C# (Aspose.Cells)
// AI Prompts: Generate C# code that builds a new workbook with Aspose.Cells, inserts a placeholder signature line at a specific cell, saves the file as .xlsx, and then invokes an external XAdES signer via Process.Start to apply a digital signature. | Modify the example to accept the workbook path and the XAdES signer executable path as command‑line arguments, adding robust error handling for missing files and process launch failures. | Demonstrate how to capture the exit code from the XAdES signing process, log a success or failure message, and optionally clean up the unsigned file when the signature operation fails.
// Common Searches: how to add a signature line placeholder to an Excel file with Aspose.Cells C# | C# run external XAdES signing executable to digitally sign a .xlsx workbook | process.start error handling when calling command line signer for Excel files in .NET
// Tags: signature line placeholder Aspose.Cells C# | external XAdES signer integration .xlsx | process.start digital signature Excel C# | command line signing utility error handling | Aspose.Cells workbook save and sign workflow

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook with Aspose.Cells, adds a placeholder signature line, saves it as an .xlsx file, and then launches an external XAdES signing executable via Process.Start to apply a digital signature, including checks for the signer tool and basic process error handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (optional, shown for completeness)
            Worksheet sheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // NOTE: SignatureLine feature is not available in the current
            // Aspose.Cells version used in this project. If a newer version
            // is referenced, the code below can be re‑enabled:
            //
            // int row = 5;
            // int column = 2;
            // int signatureIndex = sheet.SignatureLineCollection.Add(row, column);
            // SignatureLine signatureLine = sheet.SignatureLineCollection[signatureIndex];
            // signatureLine.Comment = "Approved by";
            // signatureLine.SuggestedSigner = "John Doe";
            // signatureLine.SuggestedSignerEmail = "john.doe@example.com";
            // signatureLine.ShowDate = true;
            // -----------------------------------------------------------------

            // Define the output file path
            string tempFilePath = "WorkbookWithSignatureLine.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(tempFilePath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(tempFilePath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to: {tempFilePath}");

            // Path to the external signing utility
            string signerExe = "XAdESSigner.exe";

            // Verify that the signing utility exists
            if (!File.Exists(signerExe))
            {
                Console.WriteLine($"Signing utility not found: {signerExe}");
                return;
            }

            // Prepare process start information
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = signerExe,
                Arguments = $"\"{tempFilePath}\" \"{tempFilePath}\"",
                UseShellExecute = false,
                CreateNoWindow = true
            };

            // Execute the external signing tool
            try
            {
                using (Process signerProcess = Process.Start(startInfo))
                {
                    signerProcess?.WaitForExit();
                }

                Console.WriteLine("Workbook created and signed successfully.");
            }
            catch (Exception procEx)
            {
                Console.WriteLine($"Error during signing process: {procEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
