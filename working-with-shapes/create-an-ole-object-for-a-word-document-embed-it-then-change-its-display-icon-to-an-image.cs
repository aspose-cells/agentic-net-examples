// Title: How to embed a Word .docx as an OLE object with a custom display icon in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# embed Word .docx as OLE object with custom icon | how to set DisplayAsIcon and ImageSourceFullName for OLE objects in Excel using Aspose.Cells | add OLE object with Word.Document.12 ProgID and custom icon in a .NET workbook | load icon image file for OLE object placeholder in Aspose.Cells worksheet | change OLE object label and icon path programmatically with Aspose.Cells
// Tags: embed Word document as OLE object Aspose.Cells | set custom icon for OLE object Excel .NET | DisplayAsIcon property usage Aspose.Cells | ImageSourceFullName for OLE icon Aspose.Cells | assign Word.Document.12 ProgID OLE Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectWithIconDemo
{
    // The example loads a Word .docx file and an icon image into byte arrays, creates a new workbook, adds an OLE object placeholder on the first worksheet using the icon data, assigns the Word document bytes to the object's ObjectData, enables DisplayAsIcon, sets the icon image path via ImageSourceFullName, provides a friendly label, specifies the Word.Document.12 ProgID, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the source files. Adjust these paths as needed.
            const string wordFilePath = "sample.docx";
            const string iconImagePath = "icon.png";
            const string outputWorkbookPath = "OleObjectWithIcon.xlsx";

            // Verify that the required files exist before proceeding.
            if (!File.Exists(wordFilePath))
            {
                Console.WriteLine($"Error: Word document not found at '{Path.GetFullPath(wordFilePath)}'.");
                return;
            }

            if (!File.Exists(iconImagePath))
            {
                Console.WriteLine($"Error: Icon image not found at '{Path.GetFullPath(iconImagePath)}'.");
                return;
            }

            try
            {
                // Load the Word document and the icon image into byte arrays.
                byte[] wordData = File.ReadAllBytes(wordFilePath);
                byte[] iconData = File.ReadAllBytes(iconImagePath);

                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add an OLE object placeholder using the icon image data.
                // The Add method parameters: upper left row, upper left column, width, height, image data.
                int oleIndex = sheet.OleObjects.Add(2, 2, 120, 120, iconData);
                OleObject oleObject = sheet.OleObjects[oleIndex];

                // Embed the Word document data.
                oleObject.ObjectData = wordData;

                // Configure the OLE object to display as an icon.
                oleObject.DisplayAsIcon = true;

                // Set the image source for the icon. This can be a file path or a URL.
                // Using the same icon file ensures the icon displayed matches the image we loaded.
                oleObject.ImageSourceFullName = Path.GetFullPath(iconImagePath);

                // Optional: set a friendly label that appears when the icon is hovered.
                oleObject.Label = Path.GetFileName(wordFilePath);

                // Set the ProgID for a Word document. This tells Excel which application to use.
                // For .docx files the ProgID is usually "Word.Document.12".
                oleObject.ProgID = "Word.Document.12";

                // Save the workbook.
                workbook.Save(outputWorkbookPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputWorkbookPath)}'.");
                Console.WriteLine("The OLE object embeds the Word document and displays the custom icon.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
