// Title: Add a Signature Line to the Last Worksheet in an Aspose.Cells Workbook (C#)
// Description: Creates or loads a workbook, determines the final sheet using Worksheets.Count, and inserts a configurable SignatureLine at a specific cell before saving the file.
// Keywords: Aspose.Cells C# signature line | add signature line Excel | last worksheet index | worksheet count Aspose.Cells | digital signature Excel programmatically | SignatureLine object | Aspose.Cells GitHub example
// Common Searches: Aspose.Cells add signature line to last sheet | C# get index of final worksheet Aspose.Cells | Insert digital signature line in Excel using Aspose | How to place a signature line on a specific cell with Aspose.Cells
// Developer Intent: Programmatically place a SignatureLine on the workbook's final worksheet by calculating its index with Worksheets.Count.
// Use Cases: Automatically append an approval signature to the last page of a multi‑sheet financial report. | Embed a manager’s signature line in the final sheet of a generated contract for compliance. | Add a custom signer block to the concluding worksheet of an invoice workbook.
// AI Prompts: Generate C# code that uses Aspose.Cells to add a SignatureLine to the last worksheet of an existing workbook. | Show how to set signer name, title, and email for a SignatureLine and position it at row 10, column 4 on the final sheet. | Explain how to retrieve the last worksheet index without adding extra sheets and then insert a signature line.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates or loads a workbook, determines the final sheet using Worksheets.Count, and inserts a configurable SignatureLine at a specific cell before saving the file.
class AddSignatureToLastWorksheet
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add a few worksheets for demonstration
        workbook.Worksheets.Add("FirstSheet");
        workbook.Worksheets.Add("SecondSheet");
        workbook.Worksheets.Add("ThirdSheet"); // This will be the last worksheet

        // Determine the index of the last worksheet using the worksheet count
        int lastSheetIndex = workbook.Worksheets.Count - 1;

        // Access the last worksheet
        Worksheet lastWorksheet = workbook.Worksheets[lastSheetIndex];

        // Create a SignatureLine object and set its properties
        SignatureLine signatureLine = new SignatureLine
        {
            Signer = "John Doe",
            Title = "Manager",
            Email = "john.doe@example.com",
            AllowComments = true,
            IsLine = true,
            ShowSignedDate = true,
            Instructions = "Please sign to approve."
        };

        // Add the signature line to the worksheet at a specific cell position (row 5, column 2)
        // Row and column indices are zero‑based, so row 5 = sixth row, column 2 = third column
        lastWorksheet.Shapes.AddSignatureLine(5, 2, signatureLine);

        // Save the workbook to a file
        workbook.Save("WorkbookWithSignatureLine.xlsx");
    }
}
