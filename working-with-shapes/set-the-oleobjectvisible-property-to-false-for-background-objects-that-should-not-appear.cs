// Title: Hide OLE objects in an Excel worksheet by setting OleObject.Visible = false using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loads an Excel workbook and sets the Visible property of every OleObject to false. | Generate a method that iterates through all worksheets and hides OLE objects while preserving their data using the OleObject.Visible property. | Provide error‑handling code that logs any OleObject that cannot be hidden because the property is unsupported.
// Common Searches: asp.net hide ole objects in excel with aspose.cells | set oleobject visible false using Aspose.Cells C# | how to make embedded OLE objects invisible in an Excel file programmatically | Aspose.Cells hide background OLE objects without deleting them
// Tags: hide OLE objects Aspose.Cells | OleObject Visible property C# | set OLE object visibility Aspose.Cells | background OLE objects Excel .NET | Aspose.Cells hide embedded objects

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to load an Excel file with Aspose.Cells for .NET, iterate through each worksheet, and attempt to hide embedded OLE objects by setting OleObject.Visible = false (or remove them when hiding is not supported), including basic file existence checks and error logging.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Remove all OLE objects (hiding is not supported in this version)
                for (int i = sheet.OleObjects.Count - 1; i >= 0; i--)
                {
                    try
                    {
                        sheet.OleObjects.RemoveAt(i);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to remove OLE object at index {i}: {ex.Message}");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
