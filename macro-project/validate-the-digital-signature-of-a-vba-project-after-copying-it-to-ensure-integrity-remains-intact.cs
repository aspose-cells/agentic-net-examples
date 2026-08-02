// Title: Validate VBA Project Digital Signature After Copying Workbooks with Aspose.Cells for .NET
// Description: C# example that loads a signed .xlsm file, displays the VBA project's IsSigned and IsValidSigned flags, copies the VbaProject to a new workbook, saves the result to a memory stream, reloads it, and confirms that the digital signature remains valid after the copy operation.
// Keywords: Aspose.Cells | .NET | C# | VBA project | digital signature | IsSigned | IsValidSigned | copy VBA project | preserve signature | macro-enabled workbook | memory stream | Excel automation
// Common Searches: Aspose.Cells verify VBA signature after copy | C# check VBA project IsValidSigned | preserve macro digital signature when cloning .xlsm | how to copy signed VBA project using Aspose.Cells | validate VBA digital signature in memory stream
// Developer Intent: Confirm that a VBA project's digital signature stays intact and validates after being copied to another workbook.
// Use Cases: Clone a signed Excel template while keeping its macro signature trusted. | Automate integrity checks for signed macro-enabled workbooks in a server‑side process. | Detect tampering by reloading a copied workbook and verifying the IsValidSigned flag.
// AI Prompts: Generate C# code with Aspose.Cells that copies a signed VBA project from one .xlsm file to another and returns the IsSigned and IsValidSigned values. | Show how to handle exceptions when validating VBA digital signatures after workbook manipulation using Aspose.Cells. | Create a unit test that loads a signed workbook, copies its VbaProject, saves to a memory stream, reloads, and asserts that IsValidSigned is true.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// C# example that loads a signed .xlsm file, displays the VBA project's IsSigned and IsValidSigned flags, copies the VbaProject to a new workbook, saves the result to a memory stream, reloads it, and confirms that the digital signature remains valid after the copy operation.
class ValidateVbaSignatureAfterCopy
{
    static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }

    public static void Run()
    {
        const string sourcePath = "SignedSource.xlsm";

        // Ensure the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file \"{sourcePath}\" not found.");
            return;
        }

        // Load the source workbook that already contains a signed VBA project
        Workbook sourceWb = new Workbook(sourcePath);

        // Display the signature status of the source VBA project
        Console.WriteLine("Source - IsSigned: " + sourceWb.VbaProject.IsSigned);
        Console.WriteLine("Source - IsValidSigned: " + sourceWb.VbaProject.IsValidSigned);

        // Create a new (empty) workbook that will receive the copied VBA project
        Workbook destWb = new Workbook();

        // Copy the VBA project from the source workbook to the destination workbook
        destWb.VbaProject.Copy(sourceWb.VbaProject);

        // Save the destination workbook to a memory stream
        using (MemoryStream ms = new MemoryStream())
        {
            destWb.Save(ms, SaveFormat.Xlsm);
            ms.Position = 0; // Reset stream position before reloading

            // Reload the workbook from the stream to ensure the signature is persisted
            Workbook reloadedWb = new Workbook(ms);

            // Display the signature status after copying and reloading
            Console.WriteLine("Destination after copy - IsSigned: " + reloadedWb.VbaProject.IsSigned);
            Console.WriteLine("Destination after copy - IsValidSigned: " + reloadedWb.VbaProject.IsValidSigned);
        }
    }
}
