// Title: Add a Signature Line and Password‑Protect an Excel Worksheet with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a new workbook, insert a customizable SignatureLine shape at a specific cell, protect the entire worksheet with a password using ProtectionType.All, and save the result as SignatureProtected.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# signature line | add signature line Excel Aspose | protect worksheet password Aspose.Cells | Worksheet.Protect C# | Excel digital signature Aspose | .NET Excel protection example | lock Excel sheet without edit | SignatureLine shape Aspose.Cells | Aspose.Cells worksheet security | C# Excel password protection
// Common Searches: How to insert a signature line in Excel with Aspose.Cells C# | Aspose.Cells protect worksheet with password example | C# code to lock an Excel sheet after adding a signature line | Add digital signature line and secure worksheet using Aspose.Cells | Worksheet.Protect all cells Aspose.Cells .NET
// Developer Intent: Insert a customizable signature line into a worksheet and then lock the entire sheet with a password to prevent any further edits.
// Use Cases: Create a contract workbook where a manager’s signature line is placed and the sheet is password‑protected to guarantee document integrity. | Generate a read‑only financial report that includes an approval signature line and disallows any modifications without the password. | Prepare a compliance checklist that requires a signed acknowledgment and must be locked to avoid tampering.
// AI Prompts: Generate C# code using Aspose.Cells to add a SignatureLine with custom signer details at cell B6 and protect the worksheet with a password. | Explain how Worksheet.Protect with ProtectionType.All works in Aspose.Cells and how to configure it so users can view but not edit the sheet after adding a signature line. | Show how to modify the protection settings to allow users to select locked cells while preventing any changes in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a new workbook, insert a customizable SignatureLine shape at a specific cell, protect the entire worksheet with a password using ProtectionType.All, and save the result as SignatureProtected.xlsx using Aspose.Cells for .NET.
class AddSignatureAndProtect
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the shape collection of the worksheet
        ShapeCollection shapes = worksheet.Shapes;

        // Create and configure a signature line
        SignatureLine signatureLine = new SignatureLine();
        signatureLine.AllowComments = true;
        signatureLine.Email = "example@example.com";
        signatureLine.Instructions = "Please sign to confirm.";
        signatureLine.IsLine = true;
        signatureLine.ShowSignedDate = true;
        signatureLine.Signer = "John Doe";
        signatureLine.Title = "Manager";

        // Add the signature line at row 5, column 2 (zero‑based indices)
        shapes.AddSignatureLine(5, 2, signatureLine);

        // Protect the worksheet with a password, locking all elements
        worksheet.Protect(ProtectionType.All, "MySecretPassword", null);

        // Save the workbook
        workbook.Save("SignatureProtected.xlsx");
    }
}
