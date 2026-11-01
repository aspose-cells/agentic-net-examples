// Title: Detect digitally signed worksheets and extract signer names using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to find signed worksheets in an .xlsx file | list signer names of digital signatures in Excel workbook using Aspose.Cells | use reflection to access DigitalSignatureCollection with Aspose.Cells .NET | check if an Excel workbook contains digital signatures via Aspose.Cells API | retrieve signed object information from workbook using Aspose.Cells C#
// Tags: Aspose.Cells digital signature enumeration C# | extract signer names from Excel workbook Aspose.Cells | detect signed worksheets using Aspose.Cells .NET | Aspose.Cells DigitalSignatureCollection reflection usage | C# check workbook for digital signatures Aspose.Cells

using System;
using System.IO;
using System.Collections;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, uses reflection to obtain the DigitalSignatureCollection property, verifies the presence of digital signatures, iterates through each signature, and prints the signer name and the signed object. It also includes handling for missing files, unavailable signature features, and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Attempt to retrieve the DigitalSignatureCollection via reflection
            var signatureProp = workbook.GetType().GetProperty("DigitalSignatureCollection");
            if (signatureProp == null)
            {
                Console.WriteLine("Digital signature feature is not available in this version of Aspose.Cells.");
                return;
            }

            var signaturesObj = signatureProp.GetValue(workbook);
            if (signaturesObj is IEnumerable signatures && signatures.GetEnumerator().MoveNext())
            {
                // Iterate through each digital signature using reflection to avoid compile‑time dependencies
                foreach (var signature in signatures)
                {
                    try
                    {
                        var signerProp = signature.GetType().GetProperty("Signer");
                        var signedObjectProp = signature.GetType().GetProperty("SignedObject");

                        string signerName = signerProp?.GetValue(signature)?.ToString() ?? "Unknown signer";
                        string signedObject = signedObjectProp?.GetValue(signature)?.ToString() ?? "Unknown object";

                        Console.WriteLine($"Signed Object: {signedObject}, Signer: {signerName}");
                    }
                    catch (Exception innerEx)
                    {
                        Console.WriteLine($"Failed to read a signature: {innerEx.Message}");
                    }
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain any digitally signed worksheets.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
