using System;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

class ValidateSignatureAfterPasswordChange
{
    static void Main()
    {
        // Path to the originally digitally signed workbook
        string signedPath = "SignedWorkbook.xlsx";

        // Load the signed workbook
        Workbook signedWorkbook = new Workbook(signedPath);
        Console.WriteLine("Workbook is digitally signed (original): " + signedWorkbook.IsDigitallySigned);

        // Retrieve and display the validity of each digital signature before any changes
        DigitalSignatureCollection originalSignatures = signedWorkbook.GetDigitalSignature();
        foreach (DigitalSignature sig in originalSignatures)
        {
            Console.WriteLine("Signature valid before password change: " + sig.IsValid);
        }

        // Change the workbook protection password (this modifies the file and should break the signature)
        signedWorkbook.Protect(ProtectionType.All, "newPassword");

        // Save the modified workbook to a new file
        string tamperedPath = "TamperedWorkbook.xlsx";
        signedWorkbook.Save(tamperedPath, SaveFormat.Xlsx);

        // Load the tampered workbook
        Workbook tamperedWorkbook = new Workbook(tamperedPath);
        Console.WriteLine("Workbook is digitally signed (after change): " + tamperedWorkbook.IsDigitallySigned);

        // Retrieve and display the validity of each digital signature after the password change
        DigitalSignatureCollection tamperedSignatures = tamperedWorkbook.GetDigitalSignature();
        foreach (DigitalSignature sig in tamperedSignatures)
        {
            Console.WriteLine("Signature valid after password change: " + sig.IsValid);
        }
    }
}