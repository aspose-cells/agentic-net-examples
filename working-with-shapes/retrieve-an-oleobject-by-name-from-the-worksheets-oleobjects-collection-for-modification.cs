// Title: Retrieve and modify an OleObject by its name in an Aspose.Cells worksheet using C#
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, searches the worksheet's OleObjects collection for a specific Name, and changes the object's Width, Height, and IsLocked properties. | Write a reusable function that takes inputPath, outputPath, sheetName, and oleObjectName, then updates the matching OleObject's size and lock state with Aspose.Cells.
// Common Searches: asp.net find oleobject by name in excel workbook using Aspose.Cells | change width and height of embedded OLE object in C# Aspose.Cells | unlock an OleObject programmatically in an Excel worksheet | Aspose.Cells update properties of specific OleObject on a sheet
// Tags: retrieve oleobject by name Aspose.Cells | modify oleobject size C# | set oleobject lock state Aspose.Cells | access worksheet oleobjects collection .NET | update embedded oleobject properties Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for OleObject

// // Loads 'input.xlsx', finds the OleObject named 'MyOleObject' on 'Sheet1', sets its Width to 300, Height to 200, unlocks it, and saves the workbook as 'output.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string sheetName = "Sheet1";
            const string oleObjectName = "MyOleObject";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook (lifecycle: load)
            Workbook workbook = new Workbook(inputPath);

            // Access the target worksheet
            Worksheet worksheet = workbook.Worksheets[sheetName];
            if (worksheet == null)
            {
                Console.WriteLine($"Worksheet '{sheetName}' not found.");
                return;
            }

            // Locate the OleObject by its Name property
            OleObject targetOle = null;
            foreach (OleObject obj in worksheet.OleObjects)
            {
                if (obj.Name == oleObjectName)
                {
                    targetOle = obj;
                    break;
                }
            }

            // Modify the OleObject if found
            if (targetOle != null)
            {
                targetOle.Width = 300;      // set new width
                targetOle.Height = 200;     // set new height
                targetOle.IsLocked = false; // unlock the object
            }
            else
            {
                Console.WriteLine($"OleObject named '{oleObjectName}' not found.");
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (lifecycle: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
