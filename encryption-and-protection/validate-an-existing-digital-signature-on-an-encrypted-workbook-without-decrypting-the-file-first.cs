// Title: Validate digital signatures in a password‑protected XLSX workbook using Aspose.Cells for .NET without explicit decryption
// AI Prompts: Write C# code that opens an encrypted .xlsx file with a password using Aspose.Cells LoadOptions and checks each embedded digital signature for validity. | Create a C# method that uses reflection to access the DigitalSignatureCollection of a Workbook object and prints the signer name and validation result for each signature. | Implement logic that aggregates the validation results of all signatures in an encrypted workbook and returns a single boolean indicating whether every signature is valid.
// Common Searches: asp.net how to verify digital signatures in a password protected Excel workbook using Aspose.Cells | c# load encrypted xlsx file and validate embedded digital signatures without decrypting | using reflection to access DigitalSignatureCollection in older Aspose.Cells versions | check if all signatures are valid in an encrypted workbook with Aspose.Cells .NET | aspose.cells validate workbook digital signatures when file is password protected
// Tags: aspocells validate digital signatures encrypted xlsx | c# load password protected workbook aspocells | reflection digitalsignaturecollection access aspocells | check all signatures validity encrypted excel | aspocells digital signature verification example

using System;
using System.IO;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

// The example loads a password‑protected XLSX workbook via Aspose.Cells LoadOptions, uses reflection to obtain the DigitalSignatureCollection, iterates through each signature to invoke ValidateSignature, prints signer names with validation results, and finally reports whether all signatures in the encrypted file are valid.
class ValidateSignature
{
    static void Main()
    {
        // Path to the encrypted workbook
        string filePath = "encryptedWorkbook.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File \"{filePath}\" not found.");
            return;
        }

        // Password for the encrypted workbook (required to open the file)
        string password = "yourPassword";

        try
        {
            // Load the workbook with the password; Aspose.Cells handles decryption internally.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = password };
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Use reflection to obtain the DigitalSignatureCollection (may not be present in older versions)
            var wbType = workbook.GetType();
            var sigCollectionProp = wbType.GetProperty("DigitalSignatureCollection");
            if (sigCollectionProp == null)
            {
                Console.WriteLine("Digital signature functionality is not available in this version of Aspose.Cells.");
                return;
            }

            var signaturesObj = sigCollectionProp.GetValue(workbook);
            if (signaturesObj is not IEnumerable signaturesEnum)
            {
                Console.WriteLine("No digital signatures found in the workbook.");
                return;
            }

            bool allValid = true;
            foreach (var signature in signaturesEnum)
            {
                try
                {
                    // Validate the signature via reflection
                    var validateMethod = signature.GetType().GetMethod("ValidateSignature");
                    bool isValid = validateMethod != null && (bool)validateMethod.Invoke(signature, null);

                    // Retrieve signer information via reflection
                    var signerProp = signature.GetType().GetProperty("Signer");
                    string signer = signerProp?.GetValue(signature)?.ToString() ?? "Unknown";

                    Console.WriteLine($"Signer: {signer}, Valid: {isValid}");
                    if (!isValid) allValid = false;
                }
                catch (Exception sigEx)
                {
                    Console.WriteLine($"Error validating signature: {sigEx.Message}");
                    allValid = false;
                }
            }

            Console.WriteLine($"All signatures valid: {allValid}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
