// Title: Add a Signature Line and Save Workbook with Original Format using Aspose.Cells for .NET
// Description: Loads an existing workbook, inserts a SignatureLine shape at a chosen cell, detects the source file extension, maps it to the correct SaveFormat via FileFormatUtil, and saves the signed workbook while preserving the original file type.
// Keywords: Aspose.Cells | C# signature line | Excel digital signature | preserve file format | FileFormatUtil | ExtensionToSaveFormat | save workbook same extension | add signature line C# | Aspose.Cells encryption protection | Excel workbook signing
// Common Searches: how to add a signature line to an Excel file with Aspose.Cells | save signed workbook with original .xlsx/.xls format using Aspose.Cells | convert file extension to SaveFormat enum Aspose.Cells | preserve original Excel file type when saving with Aspose.Cells | C# Aspose.Cells add digital signature placeholder
// Developer Intent: Insert a SignatureLine into a workbook and save it using the same extension and format as the source file.
// Use Cases: Create a signed copy of a compliance‑critical spreadsheet without changing its .xlsx, .xls, or .csv format. | Batch‑process multiple workbooks to add signature placeholders while keeping each file’s native format for downstream systems. | Generate signed financial or legal reports that retain the original file type to ensure compatibility with existing workflows.
// AI Prompts: Show C# code that adds a SignatureLine to the first worksheet with Aspose.Cells and saves the file preserving the original extension. | Explain how FileFormatUtil.ExtensionToSaveFormat maps extensions to SaveFormat values and how to use it when saving a workbook. | Provide an example of properly disposing the Workbook object after adding a signature line and saving the signed file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSignatureDemo
{
    // Loads an existing workbook, inserts a SignatureLine shape at a chosen cell, detects the source file extension, maps it to the correct SaveFormat via FileFormatUtil, and saves the signed workbook while preserving the original file type.
    class Program
    {
        static void Main()
        {
            // Path to the existing workbook (any supported format)
            string sourcePath = "InputWorkbook.xlsx";

            // Path where the signed workbook will be saved
            string signedPath = "SignedWorkbook.xlsx";

            // Load the workbook (uses the Workbook(string) constructor – lifecycle rule)
            Workbook workbook = new Workbook(sourcePath);

            // Add a signature line to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a SignatureLine object with default settings
            SignatureLine signatureLine = new SignatureLine();

            // Add the signature line shape at row 5, column 2 (zero‑based indices)
            // The AddSignatureLine method is part of the Shapes collection.
            sheet.Shapes.AddSignatureLine(5, 2, signatureLine);

            // Determine the original file extension to preserve format
            string originalExtension = Path.GetExtension(sourcePath); // e.g. ".xlsx"

            // Convert the extension to a SaveFormat enum value (FileFormatUtil rule)
            SaveFormat saveFormat = FileFormatUtil.ExtensionToSaveFormat(originalExtension);

            // Save the workbook using the detected SaveFormat (Save(string, SaveFormat) rule)
            workbook.Save(signedPath, saveFormat);

            // Clean up
            workbook.Dispose();

            Console.WriteLine($"Workbook signed and saved as '{signedPath}' with format {saveFormat}.");
        }
    }
}
