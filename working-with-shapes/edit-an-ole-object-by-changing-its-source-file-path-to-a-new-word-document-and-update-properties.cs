// Title: Replace linked OLE Word source and set display properties with Aspose.Cells C#
// Description: Loads an Excel workbook, scans every worksheet for linked OLE objects, changes each object's source file to a new Word document, disables auto‑update, assigns the Word ProgID, shows the object as an icon, adds a custom label, and saves the updated file.
// Keywords: Aspose.Cells OLE edit | C# change OLE source path | linked OLE Word document | set OLE ProgID | display OLE as icon | update OLE properties | Excel OLE automation
// Common Searches: how to change OLE source file in Excel using Aspose.Cells | Aspose.Cells set ProgID for Word OLE object C# | display OLE object as icon with custom label Aspose | disable auto update for linked OLE objects Aspose.Cells
// Developer Intent: Update the file path of linked OLE objects that embed Word documents and modify their visual and update settings in an Excel workbook.
// Use Cases: Re‑link all embedded Word files in a template before sending to clients. | Prevent OLE objects from refreshing automatically when the workbook opens. | Present embedded documents as icons with meaningful labels for cleaner reports.
// AI Prompts: Write C# code with Aspose.Cells that iterates through every worksheet, finds linked OLE objects, sets ObjectSourceFullName to a given Word file, disables AutoUpdate, sets ProgID to "Word.Document.12", enables DisplayAsIcon, and assigns a custom label. | Show an example that batch‑updates OLE objects in an Excel file to point to a new document path and changes their display properties using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectEditDemo
{
    // Loads an Excel workbook, scans every worksheet for linked OLE objects, changes each object's source file to a new Word document, disables auto‑update, assigns the Word ProgID, shows the object as an icon, adds a custom label, and saves the updated file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains OLE objects
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Define the new Word document path that will replace the current OLE source
            string newWordPath = @"C:\Documents\NewDocument.docx";

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all OLE objects in the current worksheet
                foreach (OleObject ole in sheet.OleObjects)
                {
                    // Update only if the OLE object is linked to a file
                    if (ole.IsLink)
                    {
                        // Change the source file path to the new Word document
                        ole.ObjectSourceFullName = newWordPath;

                        // Update additional properties as required
                        ole.AutoUpdate = false;                     // Disable automatic updates
                        ole.ProgID = "Word.Document.12";            // Set ProgID for Word documents
                        ole.DisplayAsIcon = true;                   // Show as an icon
                        ole.Label = "New Word Document";            // Icon label
                    }
                }
            }

            // Save the modified workbook
            workbook.Save("OutputWorkbook.xlsx");
        }
    }
}
