// Title: Update an OLE object in Excel to link a new Word file and modify its properties – Aspose.Cells C# example
// Description: Loads an Excel workbook, retrieves the first OLE object, changes its linked source to a new Word document with SetNativeSourceFullName and ObjectSourceFullName, then configures AutoUpdate, DisplayAsIcon, ProgID, Label, and IsLink before saving the workbook.
// Keywords: Aspose.Cells | C# | OLE object | SetNativeSourceFullName | ObjectSourceFullName | link Word document | Excel OLE edit | display as icon | ProgID | AutoUpdate | IsLink
// Common Searches: how to change OLE object source file in Excel using Aspose.Cells | Aspose.Cells C# set OLE object to display as icon | update linked Word document path for OLE object Aspose.Cells | disable auto‑update for OLE objects in Excel with Aspose | set ProgID for OLE object to Word.Document.12 Aspose.Cells
// Developer Intent: Replace the source file of an existing OLE object with a new Word document and adjust its linking and display settings using Aspose.Cells for .NET.
// Use Cases: Relink an OLE object after moving the original Word file to a different folder. | Show a linked Word document as an icon with a custom label inside an Excel sheet. | Prevent automatic refresh of a linked OLE object to keep the workbook stable.
// AI Prompts: Generate C# code that iterates over all OLE objects in a worksheet and updates each linked Word document path using Aspose.Cells. | Explain how to set ProgID, enable DisplayAsIcon, and assign a custom label for an OLE object with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectEditDemo
{
    // Loads an Excel workbook, retrieves the first OLE object, changes its linked source to a new Word document with SetNativeSourceFullName and ObjectSourceFullName, then configures AutoUpdate, DisplayAsIcon, ProgID, Label, and IsLink before saving the workbook.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains an OLE object
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one OLE object in the worksheet
            if (worksheet.OleObjects.Count == 0)
            {
                Console.WriteLine("No OLE objects found in the worksheet.");
                return;
            }

            // Get the first OLE object
            OleObject oleObject = worksheet.OleObjects[0];

            // New Word document path to link the OLE object to
            string newWordPath = @"C:\Docs\NewDocument.docx";

            // Update the native source full name (works for linked objects)
            oleObject.SetNativeSourceFullName(newWordPath);

            // Also set the ObjectSourceFullName property for completeness
            oleObject.ObjectSourceFullName = newWordPath;

            // Update additional properties as required
            oleObject.AutoUpdate = false;                 // Do not auto‑update when source changes
            oleObject.DisplayAsIcon = true;               // Show the object as an icon
            oleObject.ProgID = "Word.Document.12";        // ProgID for Word documents
            oleObject.Label = "New Word Document";        // Icon label
            oleObject.IsLink = true;                      // Ensure the object remains linked to the file

            // Save the modified workbook
            workbook.Save("output.xlsx");

            Console.WriteLine("OLE object updated and workbook saved successfully.");
        }
    }
}
