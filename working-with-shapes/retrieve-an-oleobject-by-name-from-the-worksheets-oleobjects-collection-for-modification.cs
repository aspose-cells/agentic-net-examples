// Title: How to locate and resize a named OLE object in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# find OLE object by name in a worksheet | Change size and position of an OLE object in Excel with Aspose.Cells | Iterate worksheet OleObjects collection to modify a specific OLE object | Update dimensions of a named OLE object in an .xlsx file using Aspose.Cells | C# code to adjust left, top, width, and height of an OLE object in an Excel workbook
// Tags: named OLE object retrieval Aspose.Cells | resize OLE object dimensions C# | set OLE object left top properties Aspose.Cells | worksheet OleObjects collection access | programmatic OLE object property update .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, scans the first worksheet's OleObjects collection to find an OLE object named "MyOleObject", changes its Width, Height, Left, and Top properties, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string targetOleName = "MyOleObject";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

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

            // Modify the OleObject if it exists
            if (oleObject != null)
            {
                oleObject.Width = 300;   // new width
                oleObject.Height = 200;  // new height
                oleObject.Left = 50;     // new left position
                oleObject.Top = 50;      // new top position
            }
            else
            {
                Console.WriteLine($"OleObject named '{targetOleName}' not found.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
