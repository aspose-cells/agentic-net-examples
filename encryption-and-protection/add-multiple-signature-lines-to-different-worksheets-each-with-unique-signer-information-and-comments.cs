// Title: Add Multiple Signature Lines to Different Worksheets with Aspose.Cells for .NET
// Description: Creates a new workbook, adds two extra sheets, configures three SignatureLine objects with distinct signer details, instructions, comment options, and display settings, inserts each line into a specific cell via Shapes.AddSignatureLine, and saves the file as MultipleSignatureLines.xlsx using C# and Aspose.Cells.
// Keywords: Aspose.Cells | C# signature line example | add signature line Excel | multiple worksheets signing | SignatureLine properties | digital signature .NET | Shapes.AddSignatureLine | Excel workbook protection | GitHub Aspose.Cells sample | worksheet approval workflow
// Common Searches: how to insert a signature line on a specific worksheet using Aspose.Cells | Aspose.Cells C# add signer email and comments | place digital signature in Excel cell with Aspose | multiple sheet signature line tutorial | Aspose.Cells signature line example USA | Aspose.Cells signature line guide India
// Developer Intent: Insert separate signature lines with unique signer information on several worksheets of one Excel file.
// Use Cases: Financial report where the CFO signs the summary sheet, the marketing lead signs the budget sheet, and HR signs the policy sheet. | Automated distribution of department‑specific policies that require acknowledgment and optional comments on each relevant tab. | Template generation for cross‑functional approvals, embedding a pre‑configured signature line per sheet to streamline the signing process.
// AI Prompts: Generate C# code with Aspose.Cells that adds a signature line to cell B3, allows comments, and sets custom instructions. | Show how to create three signature lines on different worksheets, each with its own signer name, title, and email, then save the workbook. | Explain how to toggle ShowSignedDate and IsLine for an existing SignatureLine object after the workbook has been written.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, adds two extra sheets, configures three SignatureLine objects with distinct signer details, instructions, comment options, and display settings, inserts each line into a specific cell via Shapes.AddSignatureLine, and saves the file as MultipleSignatureLines.xlsx using C# and Aspose.Cells.
class AddMultipleSignatureLines
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the default worksheet and add two more worksheets
        Worksheet ws1 = workbook.Worksheets[0];
        Worksheet ws2 = workbook.Worksheets.Add("Finance");
        Worksheet ws3 = workbook.Worksheets.Add("HR");

        // -------------------------------------------------
        // Signature line for the first worksheet (default)
        // -------------------------------------------------
        SignatureLine sig1 = new SignatureLine();
        sig1.Signer = "Alice Johnson";
        sig1.Title = "Chief Financial Officer";
        sig1.Email = "alice.johnson@company.com";
        sig1.Instructions = "Please sign to approve the financial report.";
        sig1.AllowComments = true;
        sig1.ShowSignedDate = true;
        sig1.IsLine = true;

        // Add the signature line at row 2, column 1 (zero‑based indices)
        ws1.Shapes.AddSignatureLine(2, 1, sig1);

        // -------------------------------------------------
        // Signature line for the second worksheet (Finance)
        // -------------------------------------------------
        SignatureLine sig2 = new SignatureLine();
        sig2.Signer = "Bob Smith";
        sig2.Title = "Head of Marketing";
        sig2.Email = "bob.smith@company.com";
        sig2.Instructions = "Sign to confirm the marketing plan.";
        sig2.AllowComments = false;
        sig2.ShowSignedDate = true;
        sig2.IsLine = true;

        // Add the signature line at row 5, column 3
        ws2.Shapes.AddSignatureLine(5, 3, sig2);

        // -------------------------------------------------
        // Signature line for the third worksheet (HR)
        // -------------------------------------------------
        SignatureLine sig3 = new SignatureLine();
        sig3.Signer = "Carol Lee";
        sig3.Title = "HR Manager";
        sig3.Email = "carol.lee@company.com";
        sig3.Instructions = "Sign to acknowledge the new policy.";
        sig3.AllowComments = true;
        sig3.ShowSignedDate = false;
        sig3.IsLine = true;

        // Add the signature line at row 1, column 0
        ws3.Shapes.AddSignatureLine(1, 0, sig3);

        // Save the workbook with all signature lines
        workbook.Save("MultipleSignatureLines.xlsx");
    }
}
