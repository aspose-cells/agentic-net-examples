// Title: Identify digitally signed worksheets in an Excel workbook and retrieve each worksheet’s signer names using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loads an .xlsx file, accesses its DigitalSignatureCollection via reflection, filters for worksheet signatures, and prints each sheet name together with its signer(s). | Create a reusable C# method that returns a Dictionary<string, List<string>> mapping worksheet names to all signer names extracted from the workbook’s digital signatures using Aspose.Cells. | Enhance the sample to handle multiple signatures per worksheet and output a distinct, comma‑separated list of signers for each signed sheet, including proper error handling for missing files.
// Common Searches: how to enumerate worksheet digital signatures in an Excel file using Aspose.Cells for .NET | C# retrieve signer name from worksheet digital signature Aspose.Cells | list all signed worksheets and their signers in a .xlsx with Aspose.Cells | using reflection to access DigitalSignatureCollection in Aspose.Cells C# example
// Tags: Aspose.Cells worksheet signature extraction | C# signer name retrieval from Excel | Excel workbook signature mapping .NET | dynamic DigitalSignatureCollection access Aspose.Cells | signed worksheet enumeration using Aspose.Cells

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, uses reflection to obtain the DigitalSignatureCollection, iterates through each signature, filters those targeting worksheets, extracts the worksheet name and signer, aggregates signers per sheet in a dictionary, and prints each signed worksheet with its associated signer names while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Dictionary to store worksheet name -> list of signer names
            var worksheetSigners = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            // Use reflection/dynamic to access digital signatures (avoids compile‑time dependency on the exact API version)
            object dscObj = null;
            var dscProp = workbook.GetType().GetProperty("DigitalSignatureCollection");
            if (dscProp != null)
            {
                dscObj = dscProp.GetValue(workbook);
            }

            if (dscObj is IEnumerable dscEnumerable)
            {
                foreach (var signature in dscEnumerable)
                {
                    // Retrieve signer name safely
                    var signerProp = signature.GetType().GetProperty("Signer");
                    string signer = signerProp?.GetValue(signature) as string ?? string.Empty;

                    // Retrieve signed object information
                    var signedInfoProp = signature.GetType().GetProperty("SignedObjectInfo");
                    var signedInfo = signedInfoProp?.GetValue(signature);
                    if (signedInfo == null) continue;

                    // Object type (e.g., Worksheet, Workbook, etc.)
                    var objectTypeProp = signedInfo.GetType().GetProperty("ObjectType");
                    var objectType = objectTypeProp?.GetValue(signedInfo)?.ToString();

                    // Process only worksheet signatures
                    if (!string.Equals(objectType, "Worksheet", StringComparison.OrdinalIgnoreCase))
                        continue;

                    // Object identifier – for worksheets this is the sheet name
                    var objectIdProp = signedInfo.GetType().GetProperty("ObjectId");
                    string sheetName = objectIdProp?.GetValue(signedInfo) as string ?? string.Empty;

                    if (string.IsNullOrEmpty(sheetName))
                        continue;

                    // Record the signer for the worksheet
                    if (!worksheetSigners.ContainsKey(sheetName))
                        worksheetSigners[sheetName] = new List<string>();

                    worksheetSigners[sheetName].Add(signer);
                }
            }

            // Output the results
            if (worksheetSigners.Count > 0)
            {
                foreach (var kvp in worksheetSigners)
                {
                    Console.WriteLine($"Worksheet '{kvp.Key}' is signed by: {string.Join(", ", kvp.Value)}");
                }
            }
            else
            {
                Console.WriteLine("No digitally signed worksheets were found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
