// Title: C# – Add a Signature Line with a Picture Placeholder and Style It using Aspose.Cells
// Description: Shows how to create a new Excel workbook with Aspose.Cells for .NET, insert an empty picture placeholder at cell B2, customize its border, size, and placement, attach a SignatureLine (signer, title, email, instructions) to the picture, and save the file as SignatureLineWithPlaceholder.xlsx.
// Keywords: Aspose.Cells | C# signature line | Excel picture placeholder | SignatureLine object | customize picture border | add signature line without image | Aspose.Cells workbook | Excel digital signature placeholder | set picture properties Aspose.Cells | signature line styling .NET
// Common Searches: Aspose.Cells add signature line to Excel | C# insert picture placeholder in Excel with Aspose.Cells | How to style a signature placeholder picture in Aspose.Cells | Attach SignatureLine to a picture using Aspose.Cells .NET | Create empty signature line in Excel workbook programmatically
// Developer Intent: Programmatically embed a signature line in an Excel sheet by linking it to a styled picture placeholder, without requiring an actual image.
// Use Cases: Create contract templates where signers see a clearly marked signature area. | Generate batch reports with individual signer placeholders for automated approval workflows. | Prepare protected Excel forms that display a placeholder until a user adds a scanned signature. | Build multi‑sheet workbooks where each sheet contains a customized signature placeholder for different stakeholders.
// AI Prompts: Generate C# code using Aspose.Cells to add an empty picture at B2, set its border color to DarkBlue, width 150, height 80, place it inside the cell, and assign a SignatureLine with signer name, title, email, and instructions. | Explain how the IsPlacedInCell and AlternativeText properties affect a picture placeholder used for a signature line in Aspose.Cells. | Show how to export the workbook to .xlsx and later replace the placeholder image with an actual signature while preserving the SignatureLine settings. | Provide a step‑by‑step guide to create multiple worksheets each containing a unique signature placeholder for different users using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSignatureLineDemo
{
    // Shows how to create a new Excel workbook with Aspose.Cells for .NET, insert an empty picture placeholder at cell B2, customize its border, size, and placement, attach a SignatureLine (signer, title, email, instructions) to the picture, and save the file as SignatureLineWithPlaceholder.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a picture placeholder (no image) at cell B2 (row 1, column 1)
            int pictureIndex = worksheet.Pictures.Add(1, 1, (string)null);
            Picture picture = worksheet.Pictures[pictureIndex];

            // Customize appearance of the placeholder
            picture.BorderLineColor = Color.DarkBlue;   // border color
            picture.BorderWeight = 2;                  // border thickness (pt)
            picture.Width = 150;                       // width in pixels
            picture.Height = 80;                       // height in pixels
            picture.IsPlacedInCell = true;             // place the picture inside the cell
            picture.AlternativeText = "Signature placeholder";

            // Create and configure the signature line
            SignatureLine signatureLine = new SignatureLine
            {
                Signer = "John Doe",
                Title = "Project Manager",
                Email = "john.doe@example.com",
                Instructions = "Please sign here to approve the document.",
                IsLine = true,
                AllowComments = true,
                ShowSignedDate = true
            };

            // Assign the signature line to the picture
            picture.SignatureLine = signatureLine;

            // Save the workbook
            workbook.Save("SignatureLineWithPlaceholder.xlsx");
        }
    }
}
