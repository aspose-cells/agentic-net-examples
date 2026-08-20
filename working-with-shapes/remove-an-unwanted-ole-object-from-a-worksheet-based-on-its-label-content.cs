// Title: Remove OLE objects by label in Excel using Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, iterates each worksheet, scans the OleObjects collection in reverse, compares each object's Label to a target string, removes matching OLE objects, and saves the updated file.
// Keywords: Aspose.Cells | C# remove OLE object | delete OLE object by label | OleObjects collection | Excel OLE removal | Aspose.Cells API | RemoveAt | batch OLE cleanup
// Common Searches: Aspose.Cells delete OLE object by label | C# remove specific OLE object from Excel | How to iterate OleObjects in Aspose.Cells | Remove unwanted OLE shapes from workbook | Batch remove OLE objects Aspose.Cells
// Developer Intent: Remove all OLE objects whose label matches a specified value from one or more worksheets.
// Use Cases: Sanitize workbooks by stripping confidential embedded documents before sharing. | Automate cleanup of placeholder OLE charts inserted during data import. | Prepare template files for production by removing development‑only OLE objects. | Integrate into CI pipelines to ensure no prohibited OLE content in generated reports.
// AI Prompts: Write C# code using Aspose.Cells to delete OLE objects with a given label from a specific worksheet. | Show how to log each removed OLE object's label and sheet name while processing a workbook. | Provide a LINQ‑based method to filter and remove OLE objects by label in Aspose.Cells. | Explain how to handle exceptions when the target label does not exist.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, iterates each worksheet, scans the OleObjects collection in reverse, compares each object's Label to a target string, removes matching OLE objects, and saves the updated file.
class RemoveOleObjectByLabel
{
    static void Main()
    {
        // Load the workbook containing OLE objects
        Workbook workbook = new Workbook("input.xlsx");

        // Define the label of the OLE object to remove
        string unwantedLabel = "UnwantedLabel";

        // Iterate through all worksheets (optional – you can target a specific sheet)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate backwards to safely remove items from the collection
            for (int i = sheet.OleObjects.Count - 1; i >= 0; i--)
            {
                OleObject ole = sheet.OleObjects[i];

                // Check the label of the OLE object
                if (ole.Label == unwantedLabel)
                {
                    // Remove the OLE object at the current index
                    sheet.OleObjects.RemoveAt(i);
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
