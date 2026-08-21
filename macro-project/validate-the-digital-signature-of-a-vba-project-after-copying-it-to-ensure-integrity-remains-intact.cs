// Title: Validate VBA Project Digital Signature After Copying – Aspose.Cells for .NET (C#)
// Description: Loads a signed macro‑enabled workbook, copies its VBA project to a new workbook, saves to a memory stream, reloads, and checks the IsSigned and IsValidSigned properties to ensure the digital signature remains intact after the copy operation.
// Keywords: Aspose.Cells | C# | VBA project signature | digital signature validation | copy VBA project | macro-enabled workbook | IsSigned | IsValidSigned | memory stream | .NET
// Common Searches: Aspose.Cells verify VBA signature after copy | C# check if VBA project stays signed | validate digital signature of copied macro workbook | IsValidSigned property usage | copy signed VBA project without losing signature
// Developer Intent: Confirm that a VBA project's digital signature stays valid after copying it to another workbook using Aspose.Cells for .NET.
// Use Cases: Migrate a signed VBA project from an existing .xlsm file to a newly generated workbook while preserving the signature. | Run automated tests that validate signature integrity without writing temporary files to disk. | Integrate signature verification into a CI/CD pipeline for macro‑enabled Excel documents.
// AI Prompts: Show C# code that copies a signed VBA project to a new workbook and verifies the signature with Aspose.Cells. | How can I use Aspose.Cells to ensure a VBA project's digital signature remains valid after copying? | Provide an example that uses memory streams to test VBA signature preservation in .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Loads a signed macro‑enabled workbook, copies its VBA project to a new workbook, saves to a memory stream, reloads, and checks the IsSigned and IsValidSigned properties to ensure the digital signature remains intact after the copy operation.
class ValidateVbaSignatureAfterCopy
{
    static void Main()
    {
        try
        {
            const string sourcePath = "SignedSource.xlsm";

            // Verify that the source workbook exists before loading.
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file '{sourcePath}' not found. Please provide a valid macro-enabled workbook.");
                return;
            }

            // Load the source workbook that already contains a signed VBA project.
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new empty workbook that will receive the copied VBA project.
            Workbook destinationWorkbook = new Workbook();

            // Ensure the destination workbook has a VBA project container.
            // Save it as a macro-enabled workbook to a memory stream and reload it.
            using (MemoryStream tempStream = new MemoryStream())
            {
                destinationWorkbook.Save(tempStream, SaveFormat.Xlsm);
                tempStream.Position = 0; // Reset stream position before reading.
                destinationWorkbook = new Workbook(tempStream);
            }

            // Copy the VBA project from the source workbook to the destination workbook.
            destinationWorkbook.VbaProject.Copy(sourceWorkbook.VbaProject);

            // Save the destination workbook to a memory stream and reload it to verify the signature.
            using (MemoryStream verificationStream = new MemoryStream())
            {
                destinationWorkbook.Save(verificationStream, SaveFormat.Xlsm);
                verificationStream.Position = 0; // Reset stream position before reading.
                Workbook verifiedWorkbook = new Workbook(verificationStream);

                // Output the signature status after copying.
                Console.WriteLine("Is VBA Project Signed after copy: " + verifiedWorkbook.VbaProject.IsSigned);
                Console.WriteLine("Is Signature Valid after copy: " + verifiedWorkbook.VbaProject.IsValidSigned);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
