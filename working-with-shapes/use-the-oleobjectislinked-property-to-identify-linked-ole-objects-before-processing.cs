// Title: C# – Detect and Update Linked OLE Objects in Excel with Aspose.Cells
// Description: Load an Excel workbook using Aspose.Cells for .NET, iterate through each worksheet’s OleObjects, use the OleObject.IsLink property to identify linked OLE objects, modify their ObjectSourceFullName (e.g., change a drive letter), skip embedded objects, and save the workbook with the updated links.
// Keywords: Aspose.Cells | C# | linked OLE objects | OleObject.IsLink | ObjectSourceFullName | update OLE link path | Excel automation | OleObject collection | skip embedded OLE | modify OLE source file
// Common Searches: How to check if an OLE object is linked using Aspose.Cells C# | Update source path of linked OLE objects in Excel with Aspose.Cells | Skip embedded OLE objects while processing a workbook in .NET | Change drive letter of OLE link in Excel programmatically | Aspose.Cells example for iterating OleObjects
// Developer Intent: Identify linked OLE objects in an Excel workbook and programmatically change their source file paths before saving.
// Use Cases: Process only linked OLE objects and ignore embedded ones across all worksheets. | Replace a specific drive letter or folder in the ObjectSourceFullName of every linked OLE object. | Automate bulk updates of OLE link locations when files are moved to a new server. | Validate that linked OLE objects point to existing files and log missing references.
// AI Prompts: Generate a C# method that returns a list of OleObject instances where IsLink is true. | Write code to change the folder path of all linked OLE objects to "D:\NewFolder" while leaving embedded objects untouched. | Provide a robust error‑handling pattern for missing source files when updating OleObject.ObjectSourceFullName with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectLinkProcessor
{
    // Load an Excel workbook using Aspose.Cells for .NET, iterate through each worksheet’s OleObjects, use the OleObject.IsLink property to identify linked OLE objects, modify their ObjectSourceFullName (e.g., change a drive letter), skip embedded objects, and save the workbook with the updated links.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each OLE object in the current worksheet
                foreach (OleObject ole in sheet.OleObjects)
                {
                    // Identify linked OLE objects using the IsLink property
                    if (!ole.IsLink)
                    {
                        // Skip processing for embedded (non‑linked) OLE objects
                        Console.WriteLine($"Worksheet \"{sheet.Name}\": Skipping embedded OLE object.");
                        continue;
                    }

                    // Process linked OLE object
                    Console.WriteLine($"Worksheet \"{sheet.Name}\": Found linked OLE object.");
                    Console.WriteLine($"Original linked file: {ole.ObjectSourceFullName}");

                    // Example modification: change the drive letter from C: to D:
                    string updatedPath = ole.ObjectSourceFullName.Replace("C:", "D:");
                    ole.ObjectSourceFullName = updatedPath;

                    Console.WriteLine($"Updated linked file: {ole.ObjectSourceFullName}");
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved with updated OLE links.");
        }
    }
}
