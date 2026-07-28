// Title: Add a Signature Line and Password‑Protect an Excel Worksheet with Aspose.Cells for .NET (C#)
// Description: Shows how to place a SignatureLine shape (signer, title, email, instructions) into a worksheet cell and then apply full worksheet protection with a password, saving the workbook as an .xlsx file.
// Keywords: Aspose.Cells | C# | signature line | Excel worksheet protection | password protect worksheet | digital signature Excel .NET | Aspose.Cells SignatureLine | lock Excel sheet | Excel security
// Common Searches: Aspose.Cells add signature line C# | How to password protect a worksheet using Aspose.Cells | Insert digital signature in Excel with Aspose.Cells .NET | C# code to lock Excel sheet after adding signature | Aspose.Cells protect worksheet without edits
// Developer Intent: Insert a signature line into a worksheet and secure the sheet with a password to prevent further edits.
// Use Cases: Create a contract workbook where a manager signs in a specific cell and the sheet becomes read‑only for all parties. | Generate a financial statement that requires a digital sign‑off before distribution, ensuring the data cannot be altered. | Automate HR onboarding forms by adding employee signature lines and locking each sheet to maintain data integrity.
// AI Prompts: Write C# code using Aspose.Cells to add a SignatureLine with custom signer details and then protect the worksheet with a given password. | Provide an example that adds multiple signature lines to different cells and applies worksheet protection while still allowing comments. | Explain how to configure ProtectionType options so users can select locked cells but cannot edit them after a signature line is added.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to place a SignatureLine shape (signer, title, email, instructions) into a worksheet cell and then apply full worksheet protection with a password, saving the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure the signature line
        SignatureLine signatureLine = new SignatureLine
        {
            Signer = "John Doe",
            Title = "Manager",
            Email = "john.doe@example.com",
            Instructions = "Please sign to confirm the content.",
            IsLine = true,
            AllowComments = true,
            ShowSignedDate = true
        };

        // Add the signature line to the worksheet (row 5, column 2 as an example)
        worksheet.Shapes.AddSignatureLine(5, 2, signatureLine);

        // Protect the worksheet with a password so no further edits are allowed
        worksheet.Protect(ProtectionType.All, "MySecretPassword", null);

        // Save the workbook
        workbook.Save("SignatureProtected.xlsx");
    }
}
