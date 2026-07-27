// Title: Add a Signature Line with Predefined Email (Contact Field) using Aspose.Cells for .NET
// Description: Demonstrates how to create a new Workbook, configure a SignatureLine with a preset signer name, title, and email address, enable line display, signed date, and comments, insert it into cell B2, and save the file as SignatureLineWithEmail.xlsx. The email is automatically written to the signature line's contact field.
// Keywords: Aspose.Cells signature line C# | predefined signer email | auto populate contact field | Excel signature line .NET | add signature line programmatically | SignatureLine Email property | Aspose.Cells workbook signing
// Common Searches: Aspose.Cells add signature line with email | set contact field for signature line C# | pre‑fill signer email in Excel using Aspose | how to auto‑populate signature line email | signature line email property Aspose.Cells
// Developer Intent: Insert a signature line into an Excel worksheet and set the signer’s email so the contact field is filled automatically.
// Use Cases: Approval sheets that record the approver’s email for audit trails. | Contract templates that embed a ready‑to‑sign line with the recipient’s contact information. | Automated reports where the signer’s name, title, and email appear without manual entry.
// AI Prompts: Generate code to add multiple signature lines, each with a different predefined email, using Aspose.Cells for .NET. | Explain how to read and verify the email address from a signature line after the workbook is signed. | Show how to customize the visual style of a signature line while keeping the contact field auto‑populated.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSignatureLineDemo
{
    // Demonstrates how to create a new Workbook, configure a SignatureLine with a preset signer name, title, and email address, enable line display, signed date, and comments, insert it into cell B2, and save the file as SignatureLineWithEmail.xlsx. The email is automatically written to the signature line's contact field.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a signature line and set its properties
            SignatureLine signatureLine = new SignatureLine
            {
                Signer = "John Doe",                 // Predefined signer name
                Title = "Approver",                  // Optional title
                Email = "john.doe@example.com",      // Predefined signer email (contact field)
                IsLine = true,                       // Display as a line
                ShowSignedDate = true,               // Show the date after signing
                AllowComments = true,                // Allow comments when signing
                Instructions = "Please sign to approve."
            };

            // Add the signature line to the worksheet at row 2, column 2 (zero‑based indices)
            worksheet.Shapes.AddSignatureLine(1, 1, signatureLine);

            // Save the workbook
            workbook.Save("SignatureLineWithEmail.xlsx");
        }
    }
}
