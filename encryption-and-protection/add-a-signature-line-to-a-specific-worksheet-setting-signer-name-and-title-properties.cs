// Title: Add a Signature Line with Signer Name & Title in Aspose.Cells for .NET
// Description: Shows how to create a workbook, get a worksheet, instantiate a SignatureLine, set its Signer and Title, insert it at a chosen cell with worksheet.Shapes.AddSignatureLine, and save the Excel file.
// Keywords: Aspose.Cells | C# | SignatureLine | AddSignatureLine | signer name | title property | Excel digital signature | worksheet shape | programmatic Excel signing | example code
// Common Searches: Aspose.Cells add signature line C# | set signer name Aspose.Cells | signature line title property .NET | insert digital signature line in Excel using Aspose | AddSignatureLine example
// Developer Intent: Insert a digital signature line into a worksheet and pre‑define the signer’s name and title.
// Use Cases: Generate contract templates that include a pre‑filled signature line for the responsible engineer. | Automate report creation with a manager’s signature line on the first sheet. | Produce compliance documents that embed signer name and title for audit trails.
// AI Prompts: Write C# code with Aspose.Cells to add a signature line to cell B5 for signer 'Jane Smith' and title 'Project Lead'. | Explain how to change the size and style of a signature line after adding it with Aspose.Cells. | Provide sample code to add multiple signature lines to different worksheets, each with unique signer information.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, get a worksheet, instantiate a SignatureLine, set its Signer and Title, insert it at a chosen cell with worksheet.Shapes.AddSignatureLine, and save the Excel file.
class AddSignatureLineExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet (or specify another index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a SignatureLine object and set signer name and title
        SignatureLine signatureLine = new SignatureLine();
        signatureLine.Signer = "John Doe";
        signatureLine.Title = "Software Developer";

        // Add the signature line to the worksheet at row 0, column 0 (top‑left cell)
        worksheet.Shapes.AddSignatureLine(0, 0, signatureLine);

        // Save the workbook to a file
        workbook.Save("SignatureLineDemo.xlsx");
    }
}
