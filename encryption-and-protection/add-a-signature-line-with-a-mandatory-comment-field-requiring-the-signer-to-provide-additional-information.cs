// Title: Add a mandatory comment signature line to an Excel worksheet with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a new Workbook, configure a SignatureLine that forces the signer to enter comments, display the signing date, set custom instructions, place the line in cell B2, and save the file as SignatureLineWithComments.xlsx using Aspose.Cells for C#.
// Keywords: Aspose.Cells signature line | C# mandatory comment signature | Excel signature line with comments | Aspose.Cells add signature line | SignatureLine AllowComments | Aspose.Cells .NET example
// Common Searches: Aspose.Cells required comment on signature line | C# add signature line with mandatory comments | How to enforce comments in Excel signature line using Aspose | Place signature line in specific cell Aspose.Cells | Show signed date with Aspose.Cells signature line
// Developer Intent: Insert a signature line into an Excel sheet that obliges the signer to provide a comment and records the signing date.
// Use Cases: Approval worksheets where each reviewer must sign and add remarks before finalization. | Contract templates that capture signer comments for audit compliance. | Automated audit logs embedding signature lines with mandatory feedback and timestamp.
// AI Prompts: Write C# code with Aspose.Cells to add a signature line that requires a comment and shows the signed date. | Explain how to check if the comment field was filled after signing an Aspose.Cells workbook. | Provide a tutorial to customize instruction text and enforce mandatory comments on a signature line in Excel.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a new Workbook, configure a SignatureLine that forces the signer to enter comments, display the signing date, set custom instructions, place the line in cell B2, and save the file as SignatureLineWithComments.xlsx using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure the signature line
        SignatureLine signatureLine = new SignatureLine();
        signatureLine.Signer = "John Doe";                     // Signer's name
        signatureLine.Title = "Approver";                      // Signer's title
        signatureLine.Email = "john.doe@example.com";          // Signer's email
        signatureLine.IsLine = true;                           // Mark as a signature line
        signatureLine.AllowComments = true;                    // Require comments from signer
        signatureLine.Instructions = "Please sign and provide comments."; // Prompt shown to signer
        signatureLine.ShowSignedDate = true;                   // Show date after signing

        // Add the signature line to the worksheet at row 2, column 2 (zero‑based indices)
        worksheet.Shapes.AddSignatureLine(1, 1, signatureLine);

        // Save the workbook with the signature line
        workbook.Save("SignatureLineWithComments.xlsx");
    }
}
