// Title: Add a Mandatory Comment Signature Line in Excel with Aspose.Cells for .NET (C#)
// Description: Shows how to create a new workbook, configure a SignatureLine with signer details, enable AllowComments to require a comment, set custom instructions, display the signed date, place the line at cell B2, and save the file as SignatureLineWithComments.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# signature line | Excel mandatory comment | AllowComments property | digital signature line | add signature line Aspose | Excel workbook signing | Aspose.Cells .NET | signature line instructions | signed date display
// Common Searches: Aspose.Cells add signature line with required comment | C# require comment on Excel signature line | How to set AllowComments in Aspose.Cells | Insert signature line at specific cell using Aspose.Cells | Aspose.Cells signature line custom instructions
// Developer Intent: Insert a signature line into an Excel worksheet that forces the signer to provide a comment before completing the signature.
// Use Cases: Contract templates where each approver must add remarks before signing. | Audit reports that need reviewers to leave comments on the signature line. | Compliance worksheets that capture the signing date together with mandatory comments. | Internal approval forms that enforce explanatory notes from signers.
// AI Prompts: Generate C# code with Aspose.Cells to add a signature line that requires a comment and shows custom instructions. | Modify the example to place the signature line at cell D5 and retrieve the signer name from a variable. | Explain the purpose of the AllowComments property and how to extract the comment after the workbook is signed. | Create a reusable method that adds a mandatory‑comment signature line to any worksheet given signer details.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a new workbook, configure a SignatureLine with signer details, enable AllowComments to require a comment, set custom instructions, display the signed date, place the line at cell B2, and save the file as SignatureLineWithComments.xlsx using Aspose.Cells for .NET.
class AddSignatureLineWithMandatoryComment
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure the signature line
        SignatureLine signatureLine = new SignatureLine
        {
            Signer = "John Doe",
            Title = "Approver",
            Email = "john.doe@example.com",
            IsLine = true,
            AllowComments = true,                     // Require comments from the signer
            Instructions = "Please add comments before signing.", // Guidance shown at signing time
            ShowSignedDate = true
        };

        // Add the signature line to the worksheet at row 2, column 2 (zero‑based indices)
        worksheet.Shapes.AddSignatureLine(1, 1, signatureLine);

        // Save the workbook
        workbook.Save("SignatureLineWithComments.xlsx");
    }
}
