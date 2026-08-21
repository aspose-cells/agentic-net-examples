// Title: Add a Signature Line to an Excel Workbook and Preserve Its Original Format with Aspose.Cells for .NET
// Description: Learn how to load an existing workbook, insert a SignatureLine shape at a chosen cell, detect the file's original extension, map it to the correct SaveFormat via FileFormatUtil, and save the workbook back without changing its format. The example includes file‑existence checks, optional signer details, and robust exception handling for .xlsx, .xls, .csv, and other supported types.
// Keywords: Aspose.Cells SignatureLine C# | preserve original file extension | FileFormatUtil ExtensionToSaveFormat | save workbook same format | add digital signature Excel | C# Excel shape insertion | Aspose.Cells save format detection | overwrite Excel file Aspose | Excel workbook protection .NET
// Common Searches: how to add a signature line to an Excel file using Aspose.Cells | save modified workbook with original extension Aspose.Cells | C# Aspose.Cells preserve file format when saving | insert SignatureLine shape in Excel with .NET | detect workbook format before saving Aspose
// Developer Intent: Insert a digital signature line into an existing Excel workbook and write the changes back using the same file type as the source.
// Use Cases: Add a SignatureLine to the first worksheet of a .xlsx file and overwrite the file while keeping the .xlsx format. | Process legacy .xls or .csv workbooks, embed a signature placeholder, and save without converting to a different format. | Validate the workbook path, optionally set signer name and email, and handle runtime errors during the save operation.
// AI Prompts: Generate C# code that loads an Excel workbook, adds a SignatureLine at cell B2, determines the original file extension, and saves the file preserving its format using Aspose.Cells. | Explain step‑by‑step how FileFormatUtil.ExtensionToSaveFormat maps file extensions to Aspose.Cells SaveFormat enums. | Create robust error‑handling logic for adding a signature line when the source file may be missing or the Aspose.Cells version lacks certain SignatureLine properties.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Learn how to load an existing workbook, insert a SignatureLine shape at a chosen cell, detect the file's original extension, map it to the correct SaveFormat via FileFormatUtil, and save the workbook back without changing its format. The example includes file‑existence checks, optional signer details, and robust exception handling for .xlsx, .xls, .csv, and other supported types.
class AddSignatureAndSave
{
    static void Main()
    {
        // Path to the existing workbook (replace with your actual file)
        string workbookPath = "SampleWorkbook.xlsx";

        // Verify that the workbook file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Workbook file not found: {workbookPath}");
            return;
        }

        try
        {
            // Load the workbook from the file
            Workbook workbook = new Workbook(workbookPath);

            // Access the first worksheet (or any worksheet you need)
            Worksheet sheet = workbook.Worksheets[0];

            // Create a signature line object
            SignatureLine signatureLine = new SignatureLine();

            // (Optional) Set additional properties on the signature line if supported
            // Note: Some older Aspose.Cells versions may not expose these properties.
            // Uncomment the lines below if your version supports them.
            // signatureLine.SuggestedSigner = "John Doe";
            // signatureLine.SuggestedSignerEmail = "john.doe@example.com";

            // Add the signature line to the worksheet at row 2, column 2 (zero‑based indexes)
            // Adjust the row/column as needed
            Picture signaturePicture = sheet.Shapes.AddSignatureLine(1, 1, signatureLine);

            // Determine the original file extension
            string extension = Path.GetExtension(workbookPath);

            // Convert the extension to the corresponding SaveFormat enum value
            SaveFormat saveFormat = FileFormatUtil.ExtensionToSaveFormat(extension);

            // Save the workbook back to the original file, preserving its format
            workbook.Save(workbookPath, saveFormat);

            Console.WriteLine("Signature line added and workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
