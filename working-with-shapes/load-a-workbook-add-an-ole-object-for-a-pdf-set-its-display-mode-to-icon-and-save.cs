using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsOlePdfExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths to required files
                string workbookPath = "input.xlsx";
                string pdfPath = "sample.pdf";
                string iconPath = "icon.png";

                // Verify that the workbook exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                // Verify that the PDF file exists
                if (!File.Exists(pdfPath))
                {
                    Console.WriteLine($"PDF file not found: {pdfPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Read the PDF data to embed
                byte[] pdfData = File.ReadAllBytes(pdfPath);

                // Read the icon image data (optional)
                byte[] iconData = null;
                if (File.Exists(iconPath))
                {
                    iconData = File.ReadAllBytes(iconPath);
                }

                // Add the OLE object (row 5, column 5, size 100x100 pixels)
                // Passing null for imageData uses the default icon
                int oleIndex = worksheet.OleObjects.Add(5, 5, 100, 100, iconData);

                // Retrieve the newly added OLE object
                OleObject oleObject = worksheet.OleObjects[oleIndex];

                // Set the ProgID for PDF files (Acrobat Reader)
                oleObject.ProgID = "AcroExch.Document.DC";

                // Embed the PDF data into the OLE object
                oleObject.ObjectData = pdfData;

                // Display the OLE object as an icon
                oleObject.DisplayAsIcon = true;

                // Optional label under the icon
                oleObject.Label = "Sample PDF";

                // Save the modified workbook
                workbook.Save("output.xlsx");
                Console.WriteLine("Workbook saved successfully as output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}