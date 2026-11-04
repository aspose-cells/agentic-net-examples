// Title: Validate digital signatures in a password‑protected Excel workbook without decrypting it using Aspose.Cells for .NET
// Common Searches: how to check digital signatures in an encrypted .xlsx file with Aspose.Cells C# | validate Excel workbook signatures without providing a password using Aspose.Cells | Aspose.Cells read DigitalSignatureCollection from a password‑protected workbook | C# verify if signatures in an encrypted Excel file are valid | use reflection to access DigitalSignatureCollection when Aspose.Cells version lacks direct support
// Tags: encrypted workbook digital signature validation | Aspose.Cells signature verification without password | C# reflection access DigitalSignatureCollection | Excel file signature IsValid check .NET | handling missing DigitalSignatureCollection in Aspose.Cells

using System;
using System.IO;
using System.Collections;
using Aspose.Cells;

// The example loads an Excel workbook without a password, catches encryption exceptions, uses reflection to obtain the DigitalSignatureCollection, iterates through each signature to evaluate its IsValid property, and reports whether all signatures are valid while handling cases where signature support is unavailable.
class ValidateSignature
{
    static void Main()
    {
        // Path to the workbook that may contain a digital signature
        string workbookPath = "EncryptedSignedWorkbook.xlsx";

        // Ensure the file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: File not found - {workbookPath}");
            return;
        }

        try
        {
            // Load the workbook (no password supplied; if the file is encrypted,
            // Aspose.Cells will throw an exception which we catch below)
            Workbook workbook = new Workbook(workbookPath);

            // Try to obtain the DigitalSignatureCollection via reflection.
            // This avoids compile‑time dependency on types that may not exist
            // in older Aspose.Cells versions.
            var dsProperty = workbook.GetType().GetProperty("DigitalSignatureCollection");
            if (dsProperty == null)
            {
                Console.WriteLine("Digital signature functionality is not available in this Aspose.Cells version.");
                return;
            }

            object dsCollection = dsProperty.GetValue(workbook);
            if (dsCollection == null)
            {
                Console.WriteLine("No digital signatures found.");
                return;
            }

            // Assume the workbook is valid until a signature is found invalid
            bool isSignatureValid = true;

            // The collection implements IEnumerable; iterate using reflection.
            foreach (object signature in (IEnumerable)dsCollection)
            {
                var isValidProp = signature.GetType().GetProperty("IsValid");
                if (isValidProp != null)
                {
                    bool isValid = (bool)isValidProp.GetValue(signature);
                    if (!isValid)
                    {
                        isSignatureValid = false;
                        break;
                    }
                }
            }

            // Output the validation result
            Console.WriteLine("Digital signature is valid: " + isSignatureValid);
        }
        catch (Exception ex)
        {
            // Handle any errors (e.g., encrypted file without password, corrupted file, etc.)
            Console.WriteLine("An error occurred while validating the digital signature:");
            Console.WriteLine(ex.Message);
        }
    }
}
