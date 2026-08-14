// Title: Refresh linked OLE objects and save the workbook with Aspose.Cells for .NET (C#)
// Description: Load an Excel file, enable AutoUpdate for every linked OLE object across all worksheets, and save the workbook so the latest external data is persisted using Aspose.Cells.
// Keywords: Aspose.Cells OLE refresh C# | linked OLE AutoUpdate | save workbook after OLE update | iterate worksheets OleObject | Excel OLE object refresh .NET | batch update linked OLE | server‑side Excel reporting Aspose
// Common Searches: how to refresh linked OLE objects in Excel with Aspose.Cells | enable AutoUpdate for OLE objects C# Aspose | save workbook after updating OLE links | iterate worksheets to update OLE objects .NET | refresh embedded charts or documents via OLE in Aspose
// Developer Intent: Programmatically turn on AutoUpdate for all linked OLE objects in a workbook and write the changes to a new file.
// Use Cases: Ensure embedded charts, Word docs, or PDFs linked as OLE reflect the latest source before distribution. | Automate a nightly job that opens multiple Excel files, refreshes linked OLE content, and saves them for downstream consumers. | Integrate OLE refresh into a web service that generates Excel reports on demand, guaranteeing up‑to‑date external data.
// AI Prompts: Write C# code that opens an Excel workbook, sets AutoUpdate on each linked OLE object, and saves the result using Aspose.Cells. | Explain the process of refreshing linked OLE objects and persisting the changes when saving a workbook with Aspose.Cells for .NET. | Create a reusable C# method that accepts an input path, refreshes all linked OLE objects, and returns the path of the saved workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Load an Excel file, enable AutoUpdate for every linked OLE object across all worksheets, and save the workbook so the latest external data is persisted using Aspose.Cells.
class RefreshOleObjects
{
    static void Main()
    {
        // Load the workbook that contains linked OLE objects
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through each OLE object in the current worksheet
            foreach (OleObject ole in sheet.OleObjects)
            {
                // Process only OLE objects that are linked to an external file
                if (ole.IsLink)
                {
                    // Enable automatic update so the linked object reflects the latest source data
                    ole.AutoUpdate = true;
                }
            }
        }

        // Save the workbook after refreshing the linked OLE objects
        workbook.Save("output.xlsx");
    }
}
