// Title: C# Example: Verify Digital Signature Becomes Invalid After Changing Workbook Protection Password with Aspose.Cells for .NET
// Description: Loads a digitally signed Excel workbook, confirms the signatures, protects the file with an initial password, changes the password, saves, and reloads to show that the workbook still reports being signed but each DigitalSignature.IsValid is false. Demonstrates how password changes break existing signatures using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | digital signature | IsValid | GetDigitalSignature | workbook protection | password change | signature invalidation | Excel XLSX | protect workbook | unprotect workbook | sample code | GitHub example | Aspose.Cells for .NET
// Common Searches: Aspose.Cells verify signature after password change | C# check Excel digital signature validity after protecting workbook | digital signature becomes invalid when workbook password is changed | Aspose.Cells .NET example for protecting and unprotecting signed workbook | how to detect broken digital signatures in Excel with Aspose.Cells
// Developer Intent: Confirm that altering the workbook protection password invalidates existing digital signatures while the workbook still reports being signed.
// Use Cases: Automated compliance scan that flags Excel files whose signatures are broken after a password update. | Unit test that asserts IsDigitallySigned remains true but DigitalSignature.IsValid returns false after re‑protecting a workbook. | Integration scenario where a system re‑encrypts signed workbooks with a new password and needs to detect signature loss.
// AI Prompts: Generate C# code using Aspose.Cells to load a signed Excel file, change its protection password, save it, and verify that all DigitalSignature.IsValid values are false. | Explain why changing the workbook protection password invalidates existing digital signatures in an Aspose.Cells workbook. | Create an MSTest unit test that ensures IsDigitallySigned stays true while each DigitalSignature.IsValid becomes false after re‑protecting the workbook with a new password.

using System;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

// Loads a digitally signed Excel workbook, confirms the signatures, protects the file with an initial password, changes the password, saves, and reloads to show that the workbook still reports being signed but each DigitalSignature.IsValid is false. Demonstrates how password changes break existing signatures using Aspose.Cells for .NET.
class ValidateSignatureAfterProtectionChange
{
    static void Main()
    {
        // Load a workbook that has already been digitally signed
        Workbook signedWorkbook = new Workbook("SignedWorkbook.xlsx");

        // Verify that the workbook contains a digital signature
        Console.WriteLine("Workbook is digitally signed (initial): " + signedWorkbook.IsDigitallySigned);

        // Retrieve the digital signatures and display their validity before any changes
        DigitalSignatureCollection originalSignatures = signedWorkbook.GetDigitalSignature();
        foreach (DigitalSignature sig in originalSignatures)
        {
            Console.WriteLine("Signature valid before protection change: " + sig.IsValid);
        }

        // Apply workbook protection with an initial password
        signedWorkbook.Protect(ProtectionType.All, "oldPassword");

        // Change the protection password: unprotect with the old password and protect again with a new one
        signedWorkbook.Unprotect("oldPassword");
        signedWorkbook.Protect(ProtectionType.All, "newPassword");

        // Save the workbook after the protection password has been changed
        string modifiedPath = "ModifiedWorkbook.xlsx";
        signedWorkbook.Save(modifiedPath, SaveFormat.Xlsx);

        // Load the modified workbook to verify the effect on the digital signature
        Workbook modifiedWorkbook = new Workbook(modifiedPath);

        // The workbook still reports that it is digitally signed (signature exists)
        Console.WriteLine("Workbook is digitally signed (after change): " + modifiedWorkbook.IsDigitallySigned);

        // Retrieve the signatures again and check their validity; they should now be invalid
        DigitalSignatureCollection modifiedSignatures = modifiedWorkbook.GetDigitalSignature();
        foreach (DigitalSignature sig in modifiedSignatures)
        {
            Console.WriteLine("Signature valid after protection change: " + sig.IsValid);
        }
    }
}
