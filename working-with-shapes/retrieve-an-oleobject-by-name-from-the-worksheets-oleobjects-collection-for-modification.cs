// Title: Retrieve and edit an OLE object by its Name in an Excel worksheet with Aspose.Cells for .NET
// AI Prompts: Search a worksheet's OleObjects collection for an object whose Name matches a target string, then change its Name, Left, and Top values using Aspose.Cells. | Load a .xlsx workbook, locate the OLE item called 'MyOleObject', rename it, reposition it, and write the updated workbook to a new file.
// Common Searches: C# Aspose.Cells find OLE object in worksheet by identifier | How to change left and top coordinates of an OLE shape in Excel via Aspose.Cells | Renaming OLE objects in a spreadsheet using Aspose.Cells API | Iterating over worksheet OleObjects to locate a specific OLE item | Saving modified workbook after updating OLE object properties with Aspose.Cells
// Tags: detect OLE shape within worksheet Aspose.Cells | set OLE object Left Top properties C# | change OLE object name Aspose.Cells | enumerate worksheet OleObjects collection .NET | persist OLE modifications with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel file, iterates the Sheet1 OleObjects collection to find an OLE object named 'MyOleObject', renames it, adjusts its Left and Top coordinates, and saves the workbook as a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string targetOleName = "MyOleObject";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the target worksheet (by name or index)
            Worksheet worksheet = workbook.Worksheets["Sheet1"]; // or workbook.Worksheets[0]

            // Locate the OleObject by its name
            OleObject oleObject = null;
            foreach (OleObject obj in worksheet.OleObjects)
            {
                if (obj.Name.Equals(targetOleName, StringComparison.OrdinalIgnoreCase))
                {
                    oleObject = obj;
                    break;
                }
            }

            if (oleObject == null)
            {
                Console.WriteLine($"OleObject named '{targetOleName}' not found in the worksheet.");
                return;
            }

            // Example modification: change the object's name and position
            oleObject.Name = "MyOleObject_Renamed";
            oleObject.Left = 100; // distance from the left edge (in points)
            oleObject.Top = 50;   // distance from the top edge (in points)

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
