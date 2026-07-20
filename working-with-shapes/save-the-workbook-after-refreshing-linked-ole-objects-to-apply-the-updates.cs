// Title: Enable AutoUpdate for Linked OLE Objects and Persist Changes with Aspose.Cells (.NET)
// Description: This example loads an Excel workbook, iterates through all worksheets, activates the AutoUpdate flag on each linked OleObject, and saves the file so external OLE sources are refreshed automatically using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | .NET | OleObject AutoUpdate | linked OLE refresh | programmatic OLE update | save workbook after OLE refresh | Excel OLE automation | batch OLE link update | Aspose.Cells API | C# OLE object handling
// Common Searches: Aspose.Cells set OleObject.AutoUpdate C# | How to refresh linked OLE objects in Excel with Aspose.Cells | Programmatically update OLE links using Aspose.Cells .NET | Save workbook after OLE refresh Aspose.Cells | Enable automatic OLE update in Excel via C#
// Developer Intent: Activate the AutoUpdate property on every linked OleObject in a workbook and write the file so the refreshed external content is stored.
// Use Cases: Prepare a multi‑sheet report by ensuring all embedded OLE charts and documents reflect the latest source data before distribution. | Integrate OLE link synchronization into an automated reporting pipeline that generates Excel files with external content. | Batch‑process a library of workbooks to turn on AutoUpdate for linked OLE objects and save them with up‑to‑date data.
// AI Prompts: Generate C# code with Aspose.Cells that loops through all worksheets, sets OleObject.AutoUpdate = true for linked objects, and saves the workbook. | Explain the effect of the OleObject.AutoUpdate property on workbook saving and how it triggers OLE link refresh. | Provide a robust Aspose.Cells example that handles missing files, read‑only workbooks, and exceptions while updating linked OLE objects.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example loads an Excel workbook, iterates through all worksheets, activates the AutoUpdate flag on each linked OleObject, and saves the file so external OLE sources are refreshed automatically using Aspose.Cells for .NET.
class RefreshOleObjectsDemo
{
    static void Main()
    {
        // Load the workbook that contains linked OLE objects
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all OLE objects on the worksheet
            foreach (OleObject ole in sheet.OleObjects)
            {
                // Process only linked OLE objects
                if (ole.IsLink)
                {
                    // Enable automatic update so the link is refreshed when the workbook is saved/opened
                    ole.AutoUpdate = true;
                }
            }
        }

        // Save the workbook – the linked OLE objects will be refreshed according to the AutoUpdate setting
        workbook.Save("output.xlsx");
    }
}
