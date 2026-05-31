using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

class Program
{
    static void Main()
    {
        try
        {
            // Verify that the source PDF exists
            const string pdfPath = "sample.pdf";
            if (!File.Exists(pdfPath))
            {
                Console.WriteLine($"Error: PDF file \"{pdfPath}\" not found.");
                return;
            }

            // Load PDF file as byte array
            byte[] pdfData = File.ReadAllBytes(pdfPath);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Placeholder image for the OLE object (1x1 transparent PNG)
            byte[] placeholderImage = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK8cAAAAASUVORK5CYII=");

            // Add an OLE placeholder at cell H4 (row index 3, column index 7) with size 200x200 pixels
            int oleIndex = worksheet.OleObjects.Add(
                topRow: 3,
                leftColumn: 7,
                height: 200,
                width: 200,
                imageData: placeholderImage);

            // Retrieve the added OleObject
            OleObject oleObject = worksheet.OleObjects[oleIndex];

            // Embed the PDF data (not linked, not displayed as an icon)
            oleObject.SetEmbeddedObject(
                linkToFile: false,
                objectData: pdfData,
                sourceFileName: pdfPath,
                displayAsIcon: false,
                label: "PDF Document");

            // Set the ProgID for PDF (helps Excel recognize the object type)
            oleObject.ProgID = "AcroExch.Document.DC";

            // Save the workbook
            const string outputPath = "OleObjectPdfDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}