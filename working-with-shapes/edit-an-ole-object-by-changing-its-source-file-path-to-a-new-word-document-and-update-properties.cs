using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class EditOleObject
{
    static void Main()
    {
        // Load an existing workbook that contains an OLE object
        Workbook workbook = new Workbook("InputWorkbook.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is at least one OLE object in the worksheet
        if (sheet.OleObjects.Count == 0)
        {
            Console.WriteLine("No OLE objects found in the worksheet.");
            return;
        }

        // Get the first OLE object
        OleObject ole = sheet.OleObjects[0];

        // New Word document path to link the OLE object to
        string newWordPath = @"C:\Documents\NewDocument.docx";

        // Change the native source file name (full path) of the OLE object
        ole.SetNativeSourceFullName(newWordPath);

        // Also update the ObjectSourceFullName property for consistency
        ole.ObjectSourceFullName = newWordPath;

        // Update additional properties as required
        ole.AutoUpdate = false;                     // Do not auto‑update when source changes
        ole.ProgID = "Word.Document.12";            // ProgID for Word documents (Office 2007+)

        // Save the modified workbook
        workbook.Save("OutputWorkbook.xlsx");

        Console.WriteLine("OLE object updated and workbook saved successfully.");
    }
}