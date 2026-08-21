// Title: Add a Signature Line to the Last Worksheet in Aspose.Cells (C#)
// Description: Creates a workbook, adds worksheets, determines the final sheet using Worksheets.Count, builds a SignatureLine object, places it at a specific cell with Shapes.AddSignatureLine, and saves the file as WorkbookWithSignatureLine.xlsx.
// Keywords: Aspose.Cells signature line C# | add signature to Excel sheet | last worksheet index Aspose.Cells | Shapes.AddSignatureLine example | programmatic Excel digital signature | C# Excel workbook protection
// Common Searches: Aspose.Cells add signature line to last sheet C# | How to get last worksheet index in Aspose.Cells | Insert signature placeholder in Excel using Aspose.Cells | C# code for signature line in Excel workbook | Aspose.Cells shape collection add signature line
// Developer Intent: Place a digital signature placeholder on the final worksheet of a workbook by calculating its index with Worksheets.Count and using Shapes.AddSignatureLine.
// Use Cases: Automatically append a manager's signature line to the closing sheet of a generated financial report. | Add a signature placeholder to the final contract worksheet before sending for approval. | Programmatically embed a signature line at a predefined cell in the last sheet of an Excel-based workflow.
// AI Prompts: Write C# code that uses Aspose.Cells to add a customizable SignatureLine to the last worksheet of an existing workbook, handling single‑sheet workbooks gracefully. | Show how to create a SignatureLine with signer details, insert it at row 5 column 2 of the final sheet, and then export the workbook to PDF. | Explain the steps to retrieve the index of the last worksheet using Worksheets.Count and add a signature line with custom properties using Shapes.AddSignatureLine.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds worksheets, determines the final sheet using Worksheets.Count, builds a SignatureLine object, places it at a specific cell with Shapes.AddSignatureLine, and saves the file as WorkbookWithSignatureLine.xlsx.
class AddSignatureToLastWorksheet
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Add a few worksheets to demonstrate locating the last one
        workbook.Worksheets.Add("FirstSheet");
        workbook.Worksheets.Add("SecondSheet");
        workbook.Worksheets.Add("ThirdSheet"); // This will be the last worksheet

        // Get the index of the last worksheet using the worksheet count
        int lastIndex = workbook.Worksheets.Count - 1;
        Worksheet lastWorksheet = workbook.Worksheets[lastIndex];

        // Create a SignatureLine object and set its properties
        SignatureLine signatureLine = new SignatureLine
        {
            Signer = "John Doe",
            Title = "Manager",
            Email = "john.doe@example.com",
            AllowComments = true,
            ShowSignedDate = true,
            IsLine = true,
            Instructions = "Please sign to approve."
        };

        // Add the signature line to the last worksheet at row 5, column 2 (zero‑based indices)
        // ShapeCollection.AddSignatureLine(topRow, leftColumn, signatureLine)
        lastWorksheet.Shapes.AddSignatureLine(5, 2, signatureLine);

        // Save the workbook (lifecycle rule: save)
        workbook.Save("WorkbookWithSignatureLine.xlsx");
    }
}
