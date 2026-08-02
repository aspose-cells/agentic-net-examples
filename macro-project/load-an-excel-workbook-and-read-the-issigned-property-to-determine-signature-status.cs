// Title: Check if an Excel workbook is digitally signed using Aspose.Cells for .NET (C#)
// Description: Load an .xlsx file with Aspose.Cells, read the Workbook.IsDigitallySigned property, and output the signature status. This example shows how to programmatically verify whether a workbook contains a digital signature in C#.
// Keywords: Aspose.Cells digital signature check | Workbook.IsDigitallySigned C# | verify Excel file signature .NET | detect signed workbook Aspose | C# read Excel digital signature | Aspose.Cells IsDigitallySigned property | Excel signature verification code | load signed workbook Aspose.Cells
// Common Searches: how to check if an Excel file is digitally signed using Aspose.Cells | C# Aspose.Cells IsDigitallySigned example | verify digital signature of .xlsx with .NET | detect signed workbook programmatically | Aspose.Cells read digital signature status
// Developer Intent: Determine programmatically whether a loaded Excel workbook contains a digital signature.
// Use Cases: Validate incoming .xlsx files before processing them in a secure workflow. | Prevent modifications to a workbook that is already signed. | Generate logs of signature status for bulk file imports.
// AI Prompts: Write a C# snippet that opens an Excel file with Aspose.Cells, checks the IsDigitallySigned flag, and catches any exceptions. | Create a script that scans a folder of .xlsx files and prints each file's digital signature state using Aspose.Cells. | Explain the steps to add a digital signature to a workbook and then verify it with Aspose.Cells in C#.

using System;
using Aspose.Cells;

// Load an .xlsx file with Aspose.Cells, read the Workbook.IsDigitallySigned property, and output the signature status. This example shows how to programmatically verify whether a workbook contains a digital signature in C#.
class CheckDigitalSignature
{
    static void Main()
    {
        // Load the Excel workbook from a file (uses the Workbook(string) constructor)
        Workbook workbook = new Workbook("SignedWorkbook.xlsx");

        // Retrieve the digital signature status via the IsDigitallySigned property
        bool isSigned = workbook.IsDigitallySigned;

        // Output the result
        Console.WriteLine("Workbook is digitally signed: " + isSigned);
    }
}
