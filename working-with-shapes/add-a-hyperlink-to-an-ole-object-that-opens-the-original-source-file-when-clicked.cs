// Title: Add a Hyperlink to an OLE Object in Aspose.Cells for .NET (C#) to Open the Source File
// Description: Demonstrates how to embed a local file as an OLE object in an Excel worksheet using Aspose.Cells, set the object to display as an icon, create a file‑URI hyperlink, assign it to the OLE object's Hyperlink.Address, and save the workbook so clicking the icon opens the original document.
// Keywords: Aspose.Cells OLE hyperlink C# | add hyperlink to OleObject | embed file as OLE Aspose.Cells | .NET Excel OLE object link | file URI hyperlink Aspose.Cells | C# Aspose.Cells example
// Common Searches: Aspose.Cells set hyperlink on OleObject | C# add OLE object with link to original file | How to open source document from OLE icon in Excel using Aspose | Create file URI for OLE object Aspose.Cells | Aspose.Cells hyperlink to local file
// Developer Intent: Create an OLE object in a worksheet and attach a file‑URI hyperlink that opens the original source document when the icon is clicked.
// Use Cases: Insert a Word document as an OLE icon in a spreadsheet with a click‑through link to the .docx file. | Embed a PDF as an OLE object and provide a hyperlink that launches the PDF from the sheet. | Add an Excel workbook as an OLE object in another workbook and link back to the source file for quick access.
// AI Prompts: Generate C# code that embeds a local file as an OLE object in Aspose.Cells and sets a file‑URI hyperlink to open the source document. | Explain how to create a file URI for a local file and assign it to an OleObject's Hyperlink.Address in Aspose.Cells for .NET. | Provide a step‑by‑step tutorial for adding an OLE object with a clickable hyperlink that opens the original file using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to embed a local file as an OLE object in an Excel worksheet using Aspose.Cells, set the object to display as an icon, create a file‑URI hyperlink, assign it to the OLE object's Hyperlink.Address, and save the workbook so clicking the icon opens the original document.
    public class OleObjectHyperlinkDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the source file that will be embedded as an OLE object
            string sourceFilePath = @"C:\Temp\SampleDocument.docx";

            // Ensure the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourceFilePath))
                throw new FileNotFoundException("Source file not found.", sourceFilePath);

            // Read the source file bytes (used as the OLE object's image/preview)
            byte[] fileBytes = File.ReadAllBytes(sourceFilePath);

            // Add the OLE object to the worksheet
            // Parameters: topRow, leftColumn, height (px), width (px), imageData
            int oleIndex = worksheet.OleObjects.Add(5, 2, 200, 300, fileBytes);

            // Retrieve the added OleObject
            OleObject oleObject = worksheet.OleObjects[oleIndex];

            // Set the OLE object to display as an icon (optional)
            oleObject.DisplayAsIcon = true;
            oleObject.Label = Path.GetFileName(sourceFilePath);

            // Create a hyperlink that points to the original source file.
            // Using the file URI scheme ensures the file opens when the hyperlink is activated.
            string fileUri = new Uri(sourceFilePath).AbsoluteUri; // e.g., "file:///C:/Temp/SampleDocument.docx"

            // Assign the hyperlink address (Hyperlink property is read‑only, so set its Address)
            Hyperlink hyperlink = oleObject.Hyperlink;
            hyperlink.Address = fileUri;

            // Save the workbook
            string outputPath = "OleObjectWithHyperlink.xlsx";
            workbook.Save(outputPath);
        }
    }
}
