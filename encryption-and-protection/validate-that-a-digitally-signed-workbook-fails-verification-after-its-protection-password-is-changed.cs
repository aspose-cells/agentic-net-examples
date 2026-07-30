// Title: Validate that changing an Excel workbook’s protection password invalidates its digital signatures using Aspose.Cells for .NET
// Description: This example loads a digitally signed Excel workbook, confirms each signature is valid, applies Workbook.Protect with a new password, saves the file, reloads it, and demonstrates that the signatures become invalid while Workbook.IsDigitallySigned still returns true.
// Keywords: Aspose.Cells | .NET | C# | digital signature verification | Excel workbook protection | password protection | signature invalidation | IsDigitallySigned | DigitalSignatureCollection | tamper detection | protect workbook | signature integrity | Aspose.Cells example | GitHub sample
// Common Searches: How to check if a signed Excel file is still valid after changing its protection password in C# | Aspose.Cells verify digital signature after workbook protection | Detect tampered digital signature in Excel using Aspose.Cells .NET | C# code to invalidate digital signatures by protecting workbook | Why does Workbook.IsDigitallySigned stay true after signature becomes invalid
// Developer Intent: Confirm that modifying the workbook’s protection password causes existing digital signatures to fail verification.
// Use Cases: Load a signed workbook and enumerate DigitalSignatureCollection to display each signature’s IsValid status before any changes. | Apply Workbook.Protect(ProtectionType.All, "NewPassword"), save the workbook, reload it, and verify that DigitalSignature.IsValid is now false. | Use Workbook.IsDigitallySigned to show the file still reports as digitally signed even though signature verification fails.
// AI Prompts: Show C# code with Aspose.Cells that verifies a digital signature becomes invalid after changing the workbook protection password. | Generate a unit test in C# asserting DigitalSignature.IsValid is true before protection and false after calling Workbook.Protect with a new password. | Explain how Aspose.Cells handles digital signature validation when workbook content is altered by protection changes.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsExamples
{
    // This example loads a digitally signed Excel workbook, confirms each signature is valid, applies Workbook.Protect with a new password, saves the file, reloads it, and demonstrates that the signatures become invalid while Workbook.IsDigitallySigned still returns true.
    public class VerifySignatureAfterProtectionChange
    {
        public static void Run()
        {
            try
            {
                // Path to the original digitally signed workbook
                string signedPath = "SignedWorkbook.xlsx";

                // Ensure the source workbook exists; create a placeholder if missing
                if (!File.Exists(signedPath))
                {
                    var placeholder = new Workbook();
                    placeholder.Worksheets[0].Cells["A1"].PutValue("Placeholder workbook");
                    placeholder.Save(signedPath, SaveFormat.Xlsx);
                }

                // Load the signed workbook
                Workbook signedWorkbook = new Workbook(signedPath);

                // Retrieve digital signatures
                DigitalSignatureCollection signatures = signedWorkbook.GetDigitalSignature();

                // Verify that the existing signatures are valid
                Console.WriteLine("Before changing protection password:");
                foreach (DigitalSignature sig in signatures)
                {
                    Console.WriteLine($"Signature IsValid = {sig.IsValid}");
                }

                // Apply workbook protection with a password (this modifies the file)
                // This will invalidate the existing digital signature because the file content changes
                signedWorkbook.Protect(ProtectionType.All, "NewProtectionPassword");

                // Save the modified workbook
                string tamperedPath = "TamperedWorkbook.xlsx";
                signedWorkbook.Save(tamperedPath, SaveFormat.Xlsx);

                // Load the tampered workbook
                Workbook tamperedWorkbook = new Workbook(tamperedPath);

                // Retrieve digital signatures again
                DigitalSignatureCollection tamperedSignatures = tamperedWorkbook.GetDigitalSignature();

                // Verify that the signatures are now invalid
                Console.WriteLine("\nAfter changing protection password:");
                foreach (DigitalSignature sig in tamperedSignatures)
                {
                    Console.WriteLine($"Signature IsValid = {sig.IsValid}");
                }

                // Demonstrate that the workbook reports being digitally signed
                Console.WriteLine($"\nWorkbook.IsDigitallySigned = {tamperedWorkbook.IsDigitallySigned}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            VerifySignatureAfterProtectionChange.Run();
        }
    }
}
