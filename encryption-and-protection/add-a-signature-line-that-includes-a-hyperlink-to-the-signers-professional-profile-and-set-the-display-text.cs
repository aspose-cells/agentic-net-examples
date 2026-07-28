// Title: Aspose.Cells for .NET – Add a Signature Line with Hyperlink to Signer's Profile (C#)
// Description: Creates a new workbook, inserts a signature line on the first worksheet, adds a cell with custom display text, attaches a hyperlink to the signer's professional profile (e.g., LinkedIn), and saves the file as an Excel workbook.
// Keywords: Aspose.Cells signature line C# | add hyperlink to Excel cell Aspose.Cells | signature line with profile link | C# Excel hyperlink display text | Aspose.Cells add signature line hyperlink
// Common Searches: how to add a signature line with a clickable profile link using Aspose.Cells | Aspose.Cells C# add hyperlink next to signature line | set custom display text for Excel hyperlink in Aspose.Cells | signature line with LinkedIn URL Aspose.Cells | add professional profile link to Excel signature line
// Developer Intent: Insert a signature line into a worksheet and link the signer’s name to a professional profile using a custom hyperlink text.
// Use Cases: Contract templates that let reviewers click a signer’s LinkedIn profile for verification. | Employee onboarding sheets with a signature line and a direct link to the staff member’s profile. | Audit reports that embed signature lines paired with profile hyperlinks for quick reviewer authentication.
// AI Prompts: Show C# code that adds a signature line and a hyperlink to the signer’s professional profile with custom display text using Aspose.Cells. | Provide an Aspose.Cells example for setting a hyperlink’s display text next to a signature line in an Excel workbook. | Explain how to configure signature line properties and attach a profile URL hyperlink in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, inserts a signature line on the first worksheet, adds a cell with custom display text, attaches a hyperlink to the signer's professional profile (e.g., LinkedIn), and saves the file as an Excel workbook.
class AddSignatureLineWithHyperlink
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Prepare the signature line object
        SignatureLine signatureLine = new SignatureLine
        {
            Signer = "John Doe",
            Title = "Senior Engineer",
            Email = "john.doe@example.com",
            IsLine = true,
            AllowComments = true,
            ShowSignedDate = true,
            // Instructions can contain any helper text; here we add a note about the profile link
            Instructions = "Please sign and review the professional profile."
        };

        // Add the signature line shape to the worksheet at row 2, column 2 (zero‑based indexes)
        Picture signaturePicture = worksheet.Shapes.AddSignatureLine(1, 1, signatureLine);

        // Add a cell that will display the hyperlink text
        int linkRow = 3;      // row index where the hyperlink will appear
        int linkColumn = 1;   // column index where the hyperlink will appear
        Cell linkCell = worksheet.Cells[linkRow, linkColumn];
        linkCell.PutValue("John Doe Professional Profile");

        // Add a hyperlink to the cell pointing to the signer's profile URL
        string profileUrl = "https://www.linkedin.com/in/johndoe";
        worksheet.Hyperlinks.Add(linkRow, linkColumn, 1, 1, profileUrl);

        // Save the workbook
        workbook.Save("SignatureLineWithHyperlink.xlsx");
    }
}
