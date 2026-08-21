// Title: C# – Retrieve an OleObject by Name and Modify Its Properties with Aspose.Cells
// Description: Load a workbook, access a worksheet, locate an OleObject whose Name matches a given string, change its Label and AutoUpdate settings, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells OleObject name lookup | C# modify OleObject label | disable OleObject auto update | iterate worksheet OleObjects collection | Aspose.Cells embedded OLE object example
// Common Searches: find OleObject by name Aspose.Cells C# | change label of embedded OLE object in Excel | set AutoUpdate false for OleObject using Aspose | C# code to edit OleObject properties in worksheet | Aspose.Cells example for OleObject manipulation
// Developer Intent: Locate a specific OleObject in a worksheet by its Name property and update its attributes programmatically.
// Use Cases: Update the display label of a linked chart embedded as an OleObject. | Turn off automatic refresh for a Word document OleObject after retrieving it by name. | Implement a batch routine that scans all OleObjects, matches names, and applies uniform property changes.
// AI Prompts: Generate C# code that uses Aspose.Cells to find an OleObject by its Name and set the Label and AutoUpdate fields. | Show how to handle the case where the requested OleObject does not exist and log a clear warning. | Create a reusable method that accepts workbook path, worksheet index, and OleObject name, then returns the modified OleObject.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Load a workbook, access a worksheet, locate an OleObject whose Name matches a given string, change its Label and AutoUpdate settings, and save the file using Aspose.Cells for .NET.
class RetrieveOleObjectByName
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Name of the OleObject to retrieve
        string targetName = "MyOleObject";

        // Access the first worksheet (adjust index as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Search for the OleObject with the specified name
        OleObject oleToModify = null;
        foreach (OleObject ole in sheet.OleObjects)
        {
            if (ole.Name == targetName)
            {
                oleToModify = ole;
                break;
            }
        }

        if (oleToModify != null)
        {
            // Example modifications
            oleToModify.Label = "UpdatedLabel";
            oleToModify.AutoUpdate = false;
        }
        else
        {
            Console.WriteLine($"OleObject with name '{targetName}' not found.");
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
