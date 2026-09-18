// Title: How to reposition an embedded OleObject in an Excel worksheet by setting Top and Left offsets using Aspose.Cells for .NET (C#)
// AI Prompts: Set the Top and Left properties of the first OleObject on a worksheet to 50 and 100 points with Aspose.Cells in C#. | Write C# code that moves an embedded OLE object to a new location by assigning integer offsets via the Aspose.Cells API. | Update an Excel file to change the position of an OleObject using Aspose.Cells' OleObject.Top and OleObject.Left members.
// Common Searches: Aspose.Cells C# change position of embedded OLE object in Excel | set oleobject top left coordinates Aspose.Cells .NET | move oleobject to specific cell location using Aspose.Cells | programmatically adjust OLE object placement in workbook with C# | how to offset OleObject location in Excel using Aspose API
// Tags: Aspose.Cells OleObject repositioning | C# set OleObject Top property | Aspose.Cells adjust OleObject coordinates | Excel OLE object placement .NET | OleObject coordinate adjustment Aspose

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an existing workbook, retrieves the first OleObject on the first worksheet, assigns new Top (50) and Left (100) offset values, and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one OleObject on the sheet
            if (sheet.OleObjects.Count == 0)
            {
                Console.WriteLine("No OleObjects found on the worksheet.");
                return;
            }

            // Retrieve the first OleObject
            OleObject oleObject = sheet.OleObjects[0];

            // Define new position (in points) and convert to integer as required by the API
            int topOffset = 50;   // distance from the top edge
            int leftOffset = 100; // distance from the left edge

            // Apply the new position
            oleObject.Top = topOffset;
            oleObject.Left = leftOffset;

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
