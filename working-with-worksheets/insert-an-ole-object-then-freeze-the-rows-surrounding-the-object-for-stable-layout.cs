// Title: How to embed an OLE object as an icon and freeze surrounding rows and columns in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: asp.net aspose.cells embed ole object and freeze panes example | c# add ole object to excel worksheet and lock top rows | how to display embedded file as an icon with aspose.cells | freeze rows above an ole object in aspose.cells workbook | set custom label for ole object in excel using aspose.cells c#
// Tags: OleObjects.Add with icon image Aspose.Cells | freeze panes based on OLE object position Aspose.Cells | embed external file as OLE object C# | display OLE object as icon Aspose.Cells | set OleObject.ObjectData property C# | custom label for OleObject Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectFreezeDemo
{
    // Demonstrates creating a new workbook, embedding an external file as an OLE object displayed as an icon with a custom label, freezing all rows and columns above and to the left of the object, and saving the result as an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the source files – adjust these paths to point to real files on your machine.
            const string imagePath = "icon.png";          // Image that will be shown as the OLE icon.
            const string oleDataPath = "sample.xlsx";    // The actual file to embed (can be any supported format).

            // Output workbook path.
            const string outputPath = "OleObjectWithFreeze.xlsx";

            try
            {
                // Verify that the required files exist.
                if (!File.Exists(imagePath))
                    throw new FileNotFoundException($"Image file not found: {imagePath}");
                if (!File.Exists(oleDataPath))
                    throw new FileNotFoundException($"OLE data file not found: {oleDataPath}");

                // Load the image and OLE data into byte arrays.
                byte[] imageBytes = File.ReadAllBytes(imagePath);
                byte[] oleBytes   = File.ReadAllBytes(oleDataPath);

                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define where the OLE object will be placed.
                // Row and column are zero‑based indexes.
                int startRow    = 5;   // Row 6 in Excel UI
                int startColumn = 2;   // Column C in Excel UI
                int width       = 150; // Width in pixels
                int height      = 120; // Height in pixels

                // Add the OLE object with the image that will be displayed.
                int oleIndex = sheet.OleObjects.Add(startRow, startColumn, width, height, imageBytes);
                OleObject oleObject = sheet.OleObjects[oleIndex];

                // Assign the embedded file data.
                oleObject.ObjectData = oleBytes;

                // Optional: set a friendly label that appears when the object is selected.
                oleObject.Label = Path.GetFileName(oleDataPath);

                // Optional: display the object as an icon (common for embedded files).
                oleObject.DisplayAsIcon = true;

                // Freeze rows and columns that are above/left of the OLE object.
                // FreezePanes(row, column, freezedRows, freezedColumns)
                sheet.FreezePanes(startRow, startColumn, startRow, startColumn);

                // Save the workbook.
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine($"OLE object inserted at row {startRow + 1}, column {startColumn + 1}.");
                Console.WriteLine($"Rows 1‑{startRow} and columns A‑{GetColumnLetter(startColumn)} are frozen.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred while creating the workbook:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }

        private static string GetColumnLetter(int columnIndex)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string result = string.Empty;
            int dividend = columnIndex + 1;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                result = letters[modulo] + result;
                dividend = (dividend - modulo) / 26;
            }

            return result;
        }
    }
}
