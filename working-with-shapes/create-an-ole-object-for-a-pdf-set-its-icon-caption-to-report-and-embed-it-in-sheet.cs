using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class EmbedPdfAsOleObject
{
    static void Main()
    {
        try
        {
            // Verify that the PDF file exists before attempting to read it
            const string pdfPath = "sample.pdf";
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException($"The PDF file '{pdfPath}' was not found.");

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Read the PDF file that will be embedded
            byte[] pdfBytes = File.ReadAllBytes(pdfPath);

            // Add an OLE object placeholder to the sheet.
            // Passing null for imageData lets Excel use the default icon.
            int oleIndex = sheet.OleObjects.Add(5, 5, 200, 200, null);

            // Retrieve the added OLE object
            OleObject ole = sheet.OleObjects[oleIndex];

            // Set the ProgID for a PDF document (helps Excel recognize the type)
            ole.ProgID = "AcroExch.Document.DC";

            // Embed the PDF data, display it as an icon, and set the icon caption to "Report"
            // Parameters: linkToFile, objectData, sourceFileName, displayAsIcon, label
            ole.SetEmbeddedObject(
                linkToFile: false,
                objectData: pdfBytes,
                sourceFileName: pdfPath,
                displayAsIcon: true,
                label: "Report"
            );

            // Save the workbook (the OLE object is now embedded in the sheet)
            const string outputPath = "WorkbookWithPdfOleObject.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}