// Title: C# – Add a Signature Line with Visible Placeholder and Style It Using Aspose.Cells
// Description: Creates a new workbook, inserts a SignatureLine at cell B2, retrieves the resulting Picture placeholder, and customizes its border color, thickness, size, cell anchoring, and alternative text before saving the file as an .xlsx document.
// Keywords: Aspose.Cells C# signature line | add signature placeholder Excel | customize signature picture Aspose | signature line border color | set picture size Aspose.Cells | place picture in cell | alternative text for Excel picture | Excel digital signature placeholder | Aspose.Cells example GitHub
// Common Searches: how to add a signature line with a visible placeholder using Aspose.Cells for .NET | change border color and thickness of a signature line picture in Excel | place signature line picture inside a specific cell with Aspose.Cells | set alternative text for a signature placeholder in C# | Aspose.Cells add signature line example GitHub
// Developer Intent: Insert a signature line with a visible placeholder image and modify its visual properties in an Excel workbook via Aspose.Cells.
// Use Cases: Add a signature line to cell B2 with signer details and display a blue‑bordered placeholder of defined dimensions. | Anchor the placeholder picture inside the target cell for precise layout control. | Provide alternative text for the placeholder to improve accessibility and searchability. | Save the customized workbook as an .xlsx file for distribution or further processing.
// AI Prompts: Write C# code that uses Aspose.Cells to add a signature line at a specified cell, retrieve the Picture object, and set its border color, weight, width, height, and placement. | Show an example of customizing the visual attributes of a signature line placeholder, including alternative text and cell anchoring, then save the workbook. | Explain how to obtain the Picture returned by Worksheet.Shapes.AddSignatureLine and modify its properties for styling a digital signature placeholder in Excel.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, inserts a SignatureLine at cell B2, retrieves the resulting Picture placeholder, and customizes its border color, thickness, size, cell anchoring, and alternative text before saving the file as an .xlsx document.
class AddSignatureLineWithPlaceholder
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a SignatureLine object and set its properties
        SignatureLine signatureLine = new SignatureLine();
        signatureLine.Signer = "John Doe";
        signatureLine.Title = "Manager";
        signatureLine.Email = "john.doe@example.com";
        signatureLine.Instructions = "Please sign here.";
        signatureLine.IsLine = true;
        signatureLine.AllowComments = true;
        signatureLine.ShowSignedDate = true;

        // Add the signature line to the worksheet at cell B2 (row index 1, column index 1)
        // This method returns the Picture that represents the signature placeholder
        Picture picture = worksheet.Shapes.AddSignatureLine(1, 1, signatureLine);

        // Customize the appearance of the picture (visible placeholder)
        picture.BorderLineColor = Color.Blue;   // Set border color
        picture.BorderWeight = 2;               // Set border thickness (points)
        picture.Width = 150;                    // Width in pixels
        picture.Height = 50;                    // Height in pixels
        picture.IsPlacedInCell = true;          // Place the picture inside the cell
        picture.AlternativeText = "Signature placeholder";

        // Save the workbook
        workbook.Save("SignatureLineWithPlaceholder.xlsx");
    }
}
