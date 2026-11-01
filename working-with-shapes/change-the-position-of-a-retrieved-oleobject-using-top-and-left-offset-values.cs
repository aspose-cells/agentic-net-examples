// Title: Shift an OleObject’s position in an Excel worksheet using top and left offsets with Aspose.Cells for .NET
// AI Prompts: Retrieve the first OleObject from a worksheet, increase its Top by 20 points and Left by 30 points, then save the workbook using Aspose.Cells. | Programmatically adjust the placement of an embedded OLE object in Excel by applying top and left offset values with C# and Aspose.Cells.
// Common Searches: C# Aspose.Cells how to move an OleObject down and right in Excel | set OleObject Top property Aspose.Cells .NET example | change position of embedded OLE object programmatically using Aspose.Cells | adjust OleObject coordinates worksheet Aspose.Cells C# tutorial | offset OleObject location in Excel file with Aspose.Cells API
// Tags: oleobject position offset Aspose.Cells | modify oleobject top left C# | adjust embedded OLE object placement .NET | excel worksheet oleobject relocation Aspose | aspocells oleobject coordinate update

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// Loads an Excel workbook, retrieves the first OleObject from the first worksheet, adds 20 points to its Top and 30 points to its Left to reposition it, and saves the updated file.
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

            // Ensure there is at least one OleObject
            if (sheet.OleObjects.Count > 0)
            {
                // Retrieve the first OleObject
                OleObject ole = sheet.OleObjects[0];

                // Apply top and left offsets
                ole.Top += 20;   // move down
                ole.Left += 30;  // move right
            }
            else
            {
                Console.WriteLine("No OleObjects found in the worksheet.");
            }

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
