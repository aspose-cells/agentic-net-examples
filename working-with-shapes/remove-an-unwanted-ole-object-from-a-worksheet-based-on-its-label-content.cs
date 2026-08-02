// Title: C# – Remove OLE Object by Label in Excel with Aspose.Cells for .NET
// Description: Loads an Excel workbook, scans the worksheet's OleObjects collection in reverse, deletes any OLE object whose Label matches a specified string, and saves the updated file.
// Keywords: Aspose.Cells | C# remove OLE object | delete OLE object by label | OleObjects collection | Aspose.Cells .NET | Excel OLE removal | programmatic OLE deletion | Aspose.Cells example | remove embedded OLE | Excel automation
// Common Searches: Aspose.Cells delete OLE object by label C# | remove specific OLE object from Excel worksheet .NET | how to iterate OleObjects and remove items Aspose.Cells | C# code to delete OLE objects in Excel | remove unwanted OLE objects using Aspose.Cells
// Developer Intent: Delete OLE objects whose label matches a given value.
// Use Cases: Strip temporary OLE charts before publishing reports | Clean legacy OLE links from generated financial workbooks | Prepare Excel files for distribution by removing embedded objects | Automate workbook sanitization in CI pipelines | Reduce file size by eliminating unnecessary OLE objects
// AI Prompts: Generate C# code using Aspose.Cells to remove all OLE objects with labels containing 'temp' from every worksheet. | Create a reusable method that takes a file path and label, removes matching OLE objects across all sheets, and returns the number removed. | Add robust error handling for missing files, empty OleObjects collection, and no matches when deleting OLE objects by label with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, scans the worksheet's OleObjects collection in reverse, deletes any OLE object whose Label matches a specified string, and saves the updated file.
class RemoveOleObjectByLabel
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Choose the worksheet to process (e.g., the first worksheet)
        Worksheet sheet = workbook.Worksheets[0];

        // The label of the OLE object that should be removed
        string unwantedLabel = "UnwantedLabel";

        // Iterate backwards because removing an item shifts subsequent indices
        for (int i = sheet.OleObjects.Count - 1; i >= 0; i--)
        {
            OleObject ole = sheet.OleObjects[i];
            if (ole.Label == unwantedLabel)
            {
                // Remove the OLE object at the current index
                sheet.OleObjects.RemoveAt(i);
                Console.WriteLine($"Removed OLE object with label '{unwantedLabel}' at index {i}.");
            }
        }

        // Save the modified workbook (replace with your desired output path)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to '{outputPath}'.");
    }
}
