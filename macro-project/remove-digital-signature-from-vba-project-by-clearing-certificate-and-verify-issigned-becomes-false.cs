// Title: C# – Remove VBA Project Digital Signature from an .xlsm Workbook with Aspose.Cells
// Description: Shows how to load a signed macro‑enabled workbook, clear its VBA digital certificate using VbaProject.Sign(null), save the file, reload it, and confirm that VbaProject.IsSigned returns false.
// Keywords: Aspose.Cells | C# VBA signature removal | VbaProject.Sign null | clear VBA digital certificate | IsSigned false | macro-enabled workbook | remove VBA project signature | programmatic signature deletion | Aspose.Cells .NET example
// Common Searches: how to delete VBA digital signature using Aspose.Cells | Aspose.Cells clear VBA project certificate C# | remove signature from .xlsm file programmatically | VbaProject.Sign null example | check if VBA project is signed Aspose.Cells
// Developer Intent: Programmatically strip the digital signature from a VBA project in an .xlsm file and verify that the project is no longer signed.
// Use Cases: Distribute a workbook without trust warnings by removing its VBA signature. | Prepare a macro‑enabled file for further editing after clearing the existing certificate. | Automate compliance checks by confirming that no VBA signatures remain after processing.
// AI Prompts: Generate C# code using Aspose.Cells to clear a VBA project's digital signature and output the IsSigned status. | Explain the effect of calling VbaProject.Sign(null) on a signed workbook and which properties change. | Provide a step‑by‑step guide to remove a VBA signature, save the workbook, and validate removal with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Shows how to load a signed macro‑enabled workbook, clear its VBA digital certificate using VbaProject.Sign(null), save the file, reload it, and confirm that VbaProject.IsSigned returns false.
class RemoveVbaSignatureDemo
{
    static void Main()
    {
        // Path to a macro‑enabled workbook that already has a signed VBA project
        string signedPath = "SignedVbaWorkbook.xlsm";

        // Load the signed workbook
        Workbook workbook = new Workbook(signedPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Show the initial signature state
        Console.WriteLine("Initially signed: " + vbaProject.IsSigned);

        // Clear the VBA signature by passing null to the Sign method
        vbaProject.Sign(null);

        // Save the workbook after removing the signature
        string unsignedPath = "UnsignedVbaWorkbook.xlsm";
        workbook.Save(unsignedPath, SaveFormat.Xlsm);

        // Reload the saved file to verify that the signature has been removed
        Workbook reloaded = new Workbook(unsignedPath);
        Console.WriteLine("After removal, signed: " + reloaded.VbaProject.IsSigned);
    }
}
