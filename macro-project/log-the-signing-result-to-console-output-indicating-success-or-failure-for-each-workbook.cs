// Title: Log digital signature outcome for multiple Excel workbooks using Aspose.Cells (.NET)
// Description: C# sample that loads a PFX certificate, iterates over a set of .xlsx files, adds a digital signature via reflection (fallback for older Aspose.Cells versions), saves a signed copy, reloads it to read the IsDigitallySigned flag, and writes a concise console message indicating success or failure for each workbook.
// Keywords: Aspose.Cells digital signature | C# Excel signing log | IsDigitallySigned check | multiple workbook signing | reflection fallback Aspose.Cells | certificate based Excel signature | console output signing result
// Common Searches: how to log digital signature status for each Excel file using Aspose.Cells | C# verify if workbook is digitally signed after saving | add DigitalSignatureCollection with reflection Aspose.Cells | console message for Excel signing success or failure | sign multiple Excel workbooks .NET certificate
// Developer Intent: Display in the console whether each processed workbook was signed successfully or not.
// Use Cases: Batch‑sign a collection of Excel files with a PFX certificate and track the result per file. | Provide clear error messages when the certificate or a workbook is missing. | Support older Aspose.Cells releases by using reflection to access DigitalSignatureCollection. | Confirm the signature by reloading the saved file and checking the IsDigitallySigned property.
// AI Prompts: Write C# code that signs several Excel workbooks with a PFX certificate using Aspose.Cells and logs success or failure to the console. | Show robust error handling for missing certificates, absent files, and API version differences when adding a digital signature. | Demonstrate how to reload a saved workbook, read the IsDigitallySigned flag, and output a concise verification message.

using System;
using System.IO;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsDigitalSignatureLogging
{
    // C# sample that loads a PFX certificate, iterates over a set of .xlsx files, adds a digital signature via reflection (fallback for older Aspose.Cells versions), saves a signed copy, reloads it to read the IsDigitallySigned flag, and writes a concise console message indicating success or failure for each workbook.
    class Program
    {
        static void Main()
        {
            // Paths of workbooks to be signed
            string[] workbookPaths = new string[]
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
            };

            // Certificate used for signing (replace with actual path and password)
            string certificatePath = @"certs\mycert.pfx";
            string certificatePassword = "password";

            X509Certificate2 certificate = null;

            // Load the certificate safely
            try
            {
                if (!File.Exists(certificatePath))
                {
                    Console.WriteLine($"Certificate file not found: {certificatePath}");
                }
                else
                {
                    // Load certificate (obsolete warning suppressed; still functional)
                    certificate = new X509Certificate2(certificatePath, certificatePassword, X509KeyStorageFlags.MachineKeySet);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load certificate: {ex.Message}");
            }

            if (certificate == null)
            {
                Console.WriteLine("Signing process aborted due to missing or invalid certificate.");
                return;
            }

            foreach (string inputPath in workbookPaths)
            {
                try
                {
                    // Verify workbook file exists
                    if (!File.Exists(inputPath))
                    {
                        Console.WriteLine($"Workbook file not found: {inputPath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(inputPath);

                    // Create a digital signature instance
                    DigitalSignature signature = new DigitalSignature(certificate, "Signed by Aspose.Cells", DateTime.Now);

                    // Attempt to add the signature using reflection (covers versions where the API exists)
                    PropertyInfo dsCollectionProp = workbook.GetType().GetProperty("DigitalSignatureCollection");
                    if (dsCollectionProp != null)
                    {
                        object dsCollection = dsCollectionProp.GetValue(workbook);
                        MethodInfo addMethod = dsCollection?.GetType().GetMethod("Add");
                        addMethod?.Invoke(dsCollection, new object[] { signature });
                    }
                    else
                    {
                        Console.WriteLine("DigitalSignatureCollection property not available in this Aspose.Cells version. Skipping actual signing.");
                    }

                    // Save the (potentially) signed workbook (creates a new file)
                    string signedPath = Path.Combine(
                        Path.GetDirectoryName(inputPath) ?? string.Empty,
                        Path.GetFileNameWithoutExtension(inputPath) + "_Signed" + Path.GetExtension(inputPath));

                    workbook.Save(signedPath);

                    // Reload the saved file to verify the signature status (if supported)
                    Workbook signedWorkbook = new Workbook(signedPath);
                    bool isSigned = signedWorkbook.IsDigitallySigned;

                    // Log the result
                    Console.WriteLine($"Workbook: {Path.GetFileName(inputPath)} - Signing {(isSigned ? "succeeded" : "failed")}.");
                }
                catch (Exception ex)
                {
                    // Log any exception as a failure
                    Console.WriteLine($"Workbook: {Path.GetFileName(inputPath)} - Signing failed with error: {ex.Message}");
                }
            }
        }
    }
}
