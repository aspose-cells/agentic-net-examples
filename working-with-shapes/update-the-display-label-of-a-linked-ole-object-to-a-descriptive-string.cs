// Title: C# – Set a Custom Display Label for a Linked OLE Object in Excel with Aspose.Cells
// Description: Demonstrates how to add a linked OLE object (e.g., a Word document) to an Excel worksheet, assign a meaningful label via the OleObject.Label property, save the workbook, and verify that the label persists after reloading.
// Keywords: Aspose.Cells C# OLE object label | set linked OLE object caption .NET | Excel OleObject.Label property | custom display name for OLE in Excel | programmatically change OLE object label
// Common Searches: change OLE object label Aspose.Cells | set custom caption for linked OLE object C# | update display text of OLE object in generated Excel | how to modify OleObject.Label in .NET
// Developer Intent: Assign a readable label to a linked OLE object in an Excel file.
// Use Cases: Provide a clear, user‑friendly name for a linked Word file in an automated report. | Differentiate multiple linked OLE objects by giving each a distinct caption. | Confirm that the custom label is stored correctly after the workbook is saved and reopened.
// AI Prompts: Show C# code that adds a linked OLE object to a worksheet and sets its Label property to a custom string using Aspose.Cells. | Generate an example that saves an Excel file with a labeled OLE object and then reads back the label to verify it. | Explain how to retrieve and update the Label of an existing OleObject in a loaded workbook with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a linked OLE object (e.g., a Word document) to an Excel worksheet, assign a meaningful label via the OleObject.Label property, save the workbook, and verify that the label persists after reloading.
class UpdateOleLabel
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the file that will be linked as an OLE object
            string sourceFilePath = "sample.docx";

            // Ensure the source file exists (create an empty file if necessary)
            if (!File.Exists(sourceFilePath))
            {
                File.WriteAllBytes(sourceFilePath, new byte[0]);
            }

            // Add an OLE object placeholder to the worksheet.
            // Passing null for imageData lets Aspose.Cells use the default icon.
            int oleIndex = worksheet.OleObjects.Add(5, 5, 150, 150, null);
            OleObject oleObject = worksheet.OleObjects[oleIndex];

            // Configure the OLE object as a linked object
            oleObject.IsLink = true; // Mark as linked
            oleObject.ObjectSourceFullName = Path.GetFullPath(sourceFilePath); // Set source file path
            oleObject.ProgID = "Word.Document.12"; // ProgID for a Word document

            // Update the display label of the linked OLE object
            oleObject.Label = "Project Specification";

            // Save the workbook
            string outputPath = "OleLabelDemo.xlsx";
            workbook.Save(outputPath);

            // Reload the workbook to verify that the label was saved correctly
            if (File.Exists(outputPath))
            {
                Workbook loadedWorkbook = new Workbook(outputPath);
                OleObject loadedOle = loadedWorkbook.Worksheets[0].OleObjects[0];
                Console.WriteLine("Loaded OLE Object Label: " + loadedOle.Label);
            }
            else
            {
                Console.WriteLine("Failed to save the workbook.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
