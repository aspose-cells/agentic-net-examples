// Title: Validate that changing a workbook’s protection password invalidates its digital signature with Aspose.Cells for .NET
// Description: This C# example loads a digitally signed Excel workbook, confirms the signature is present and valid, applies a new protection password via Workbook.Protect, saves the file, reloads it, and shows that while IsDigitallySigned stays true, each DigitalSignature.IsValid becomes false, indicating the signature was broken by the password change.
// Keywords: Aspose.Cells | C# digital signature verification | Excel workbook protection password | DigitalSignature.IsValid false | tampered signed workbook | Workbook.Protect | digital signature invalidation | Aspose.Cells .NET | Excel file tampering detection | region:US | region:EU
// Common Searches: Aspose.Cells verify digital signature after changing protection password | C# check if Excel signature is broken by Workbook.Protect | DigitalSignature.IsValid false after workbook tampering | How to detect altered signed Excel file using Aspose.Cells | Validate Excel digital signature integrity in .NET
// Developer Intent: Ensure that modifying the protection password of a signed workbook renders its digital signature invalid.
// Use Cases: Audit signed Excel files for unauthorized protection changes | Automated pipeline to flag tampered workbooks by checking DigitalSignature.IsValid | Unit testing of signature integrity before and after workbook protection updates | Compliance reporting for documents whose signatures become invalid after password modification
// AI Prompts: Generate C# code using Aspose.Cells that loads a signed workbook, changes its protection password, saves it, and verifies DigitalSignature.IsValid is false. | Explain why Workbook.Protect breaks an existing digital signature in an Excel file when using Aspose.Cells. | Create a NUnit test that asserts DigitalSignature.IsValid is true for the original file and false after applying a new password. | Provide a step‑by‑step guide to detect tampered signed workbooks in a .NET application with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsDigitalSignatureVerification
{
    // This C# example loads a digitally signed Excel workbook, confirms the signature is present and valid, applies a new protection password via Workbook.Protect, saves the file, reloads it, and shows that while IsDigitallySigned stays true, each DigitalSignature.IsValid becomes false, indicating the signature was broken by the password change.
    class Program
    {
        static void Main()
        {
            // Path to the original digitally signed workbook
            string signedPath = "SignedWorkbook.xlsx";

            // Load the signed workbook (create/load rule)
            Workbook signedWorkbook = new Workbook(signedPath);

            // Verify that the workbook reports being digitally signed
            bool isSigned = signedWorkbook.IsDigitallySigned;
            Console.WriteLine("Original workbook is digitally signed: " + isSigned);

            // Retrieve the digital signatures and check their validity
            DigitalSignatureCollection originalSignatures = signedWorkbook.GetDigitalSignature();
            foreach (DigitalSignature sig in originalSignatures)
            {
                Console.WriteLine("Original signature IsValid: " + sig.IsValid);
            }

            // Change the workbook protection password (this modifies the file and should break the signature)
            signedWorkbook.Protect(ProtectionType.All, "NewPassword123");

            // Save the modified workbook (save rule)
            string tamperedPath = "TamperedWorkbook.xlsx";
            signedWorkbook.Save(tamperedPath);

            // Load the tampered workbook
            Workbook tamperedWorkbook = new Workbook(tamperedPath);

            // The workbook still reports being digitally signed, but the signature should now be invalid
            bool isStillSigned = tamperedWorkbook.IsDigitallySigned;
            Console.WriteLine("Tampered workbook reports digitally signed: " + isStillSigned);

            // Retrieve signatures from the tampered workbook and verify validity
            DigitalSignatureCollection tamperedSignatures = tamperedWorkbook.GetDigitalSignature();
            foreach (DigitalSignature sig in tamperedSignatures)
            {
                Console.WriteLine("Tampered signature IsValid (should be false): " + sig.IsValid);
            }
        }
    }
}
