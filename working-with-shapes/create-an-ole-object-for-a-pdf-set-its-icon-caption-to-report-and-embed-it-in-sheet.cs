// Title: C# – Embed a PDF as an OLE object with a custom “Report” icon caption using Aspose.Cells
// Description: Creates a workbook, reads a PDF file, adds an OLE placeholder, embeds the PDF as an embedded object displayed as an icon, sets the icon label to "Report", and saves the result as XLSX. Shows how to use Aspose.Cells SetEmbeddedObject parameters in .NET.
// Keywords: Aspose.Cells | C# | embed PDF | OLE object | icon caption | DisplayAsIcon | SetEmbeddedObject | Excel automation | worksheet OLE | PDF icon in Excel
// Common Searches: embed pdf ole object aspose.cells c# | set custom icon label for ole object aspose.cells | display pdf as icon in excel using aspose | c# code to add pdf ole object with caption | asp.net embed pdf in worksheet as icon
// Developer Intent: Add a PDF to an Excel worksheet as an embedded OLE object, show it as an icon, and assign the label "Report" with Aspose.Cells for .NET.
// Use Cases: Attach a detailed PDF analysis to a financial summary workbook, keeping the sheet tidy with a clickable "Report" icon. | Build a template that stores supporting documentation (PDF) inside the spreadsheet, using an icon to preserve layout. | Automate the consolidation of multiple PDFs into a single workbook, each represented by a custom‑labeled icon for quick access.
// AI Prompts: Generate C# code with Aspose.Cells that inserts a PDF as an OLE object, displays it as an icon, and sets the icon caption to "Report" at a specific cell range. | Explain the parameters of SetEmbeddedObject when embedding a PDF as an OLE object with Aspose.Cells for .NET. | Provide a step‑by‑step tutorial for embedding a PDF as an OLE object, ensuring it is embedded (not linked) and customizing the icon label using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, reads a PDF file, adds an OLE placeholder, embeds the PDF as an embedded object displayed as an icon, sets the icon label to "Report", and saves the result as XLSX. Shows how to use Aspose.Cells SetEmbeddedObject parameters in .NET.
class EmbedPdfOleObject
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the PDF file to embed
            string pdfPath = "sample.pdf";

            // Read PDF file bytes if the file exists; otherwise use an empty array
            byte[] pdfData;
            if (File.Exists(pdfPath))
            {
                pdfData = File.ReadAllBytes(pdfPath);
            }
            else
            {
                Console.WriteLine($"PDF file not found: {pdfPath}");
                pdfData = new byte[0];
            }

            // Add an OLE object placeholder to the sheet.
            // Image data is set to an empty byte array because we will display it as an icon.
            int oleIndex = sheet.OleObjects.Add(5, 2, 150, 150, new byte[0]);

            // Retrieve the added OleObject
            OleObject ole = sheet.OleObjects[oleIndex];

            // Embed the PDF data, display it as an icon, and set the icon caption to "Report"
            // linkToFile = false (embed the file), displayAsIcon = true, label = "Report"
            ole.SetEmbeddedObject(false, pdfData, Path.GetFileName(pdfPath), true, "Report");

            // Ensure the object is shown as an icon (redundant but explicit)
            ole.DisplayAsIcon = true;

            // Save the workbook
            workbook.Save("OleObjectPdfReport.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
