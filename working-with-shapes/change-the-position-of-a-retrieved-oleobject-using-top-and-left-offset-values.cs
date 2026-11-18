// Title: Move an embedded OLE object to a new top‑left position in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: how to set OleObject top and left values with Aspose.Cells C# | Aspose.Cells .NET move embedded OLE object to specific coordinates | programmatically change position of OLE object in Excel workbook using Aspose | adjust OLE object offset in worksheet with Aspose.Cells API | C# code to reposition OleObject in Excel file using Aspose.Cells
// Tags: Aspose.Cells set OleObject position | OleObject Top property .NET | adjust OLE object coordinates Excel | move embedded OLE object Aspose.Cells | change OleObject offset C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook, accesses the first worksheet, checks for an embedded OLE object, retrieves it, updates its Top and Left properties with specified offset values, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one OLE object
            if (sheet.OleObjects.Count == 0)
            {
                Console.WriteLine("No OLE objects found in the worksheet.");
                return;
            }

            // Retrieve the first OLE object
            OleObject oleObject = sheet.OleObjects[0];

            // Define new top and left offset values (in points)
            int newTopOffset = 50;   // example top offset
            int newLeftOffset = 100; // example left offset

            // Change the position of the OLE object
            oleObject.Top = newTopOffset;
            oleObject.Left = newLeftOffset;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
