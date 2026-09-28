// Title: Add a second digital signature to an already signed Excel workbook and verify both signatures with Aspose.Cells for .NET
// AI Prompts: Load an existing signed .xlsx file, append a second PFX certificate signature using Aspose.Cells, and save the workbook. | Reopen the saved workbook and enumerate its signature collection to confirm that two signatures are present.
// Common Searches: how to append another digital signature to a signed Excel file using Aspose.Cells .NET | Aspose.Cells verify multiple signatures in an Excel workbook | C# add second certificate signature to existing signed .xlsx with Aspose.Cells | list all digital signatures in an Excel workbook after adding a new one using Aspose.Cells
// Tags: add digital signature Aspose.Cells .NET | append second certificate signature Excel | verify multiple signatures Aspose.Cells | signature collection enumeration Excel workbook | digital signature persistence Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSignatureDemo
{
    // The example loads a previously signed Excel workbook, accesses its signature collection via reflection, adds a second digital signature from a PFX certificate, saves the file, reloads it, and enumerates the signatures to display the total count and details, confirming that both signatures persist.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "signed.xlsx";
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the already signed Excel file
                Workbook workbook = new Workbook(inputPath);

                // Path to the second digital certificate and its password
                const string secondCertPath = "cert2.pfx";
                const string secondCertPassword = "password2";

                if (!File.Exists(secondCertPath))
                {
                    Console.WriteLine($"Certificate file '{secondCertPath}' not found.");
                    return;
                }

                // Obtain the signature collection (property name may vary by Aspose.Cells version)
                var sigProp = workbook.GetType().GetProperty("Signatures") ??
                              workbook.GetType().GetProperty("SignatureCollection");

                if (sigProp == null)
                {
                    Console.WriteLine("Signature collection property not found in the current Aspose.Cells version.");
                    return;
                }

                dynamic signatures = sigProp.GetValue(workbook);

                // Add the second digital signature
                signatures.Add(
                    secondCertPath,
                    secondCertPassword,
                    "Second signature",          // Reason
                    "Office",                    // Location
                    "john.doe@example.com",      // ContactInfo
                    DateTime.Now                 // SigningTime
                );

                // Save the workbook containing both signatures
                const string outputPath = "signed_with_two_signatures.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");

                // Reload the saved file to verify that both signatures persist
                if (!File.Exists(outputPath))
                {
                    Console.WriteLine($"Saved file '{outputPath}' not found.");
                    return;
                }

                Workbook verificationWorkbook = new Workbook(outputPath);
                var verSigProp = verificationWorkbook.GetType().GetProperty("Signatures") ??
                                 verificationWorkbook.GetType().GetProperty("SignatureCollection");

                if (verSigProp == null)
                {
                    Console.WriteLine("Signature collection property not found during verification.");
                    return;
                }

                dynamic verificationSignatures = verSigProp.GetValue(verificationWorkbook);

                // Output the total number of signatures
                Console.WriteLine("Total signatures: " + verificationSignatures.Count);

                // List details of each signature
                for (int i = 0; i < verificationSignatures.Count; i++)
                {
                    var signature = verificationSignatures[i];
                    Console.WriteLine(
                        $"Signature {i + 1}: Reason = {signature.Reason}, Location = {signature.Location}, Contact = {signature.ContactInfo}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
