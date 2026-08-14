// Title: Detect and Update Linked OLE Objects in Excel with Aspose.Cells for .NET
// Description: Shows how to load a workbook, iterate worksheets, filter OLE objects using OleObject.IsLink, read and change the ObjectSourceFullName of linked OLE items, and save the updated file.
// Keywords: Aspose.Cells | C# | .NET | OleObject.IsLink | linked OLE objects | ObjectSourceFullName | modify OLE path | Excel OLE handling | workbook automation
// Common Searches: filter linked OLE objects Aspose.Cells | change OLE source path C# Excel | skip embedded OLE objects Aspose | OleObject.IsLink example | update external OLE links in workbook
// Developer Intent: Identify only linked OLE objects, adjust their source file paths, and save the workbook with the corrected references.
// Use Cases: Exclude embedded OLE objects from processing to improve performance. | Rewrite the drive letter or folder in ObjectSourceFullName for all linked OLE items. | Persist the modified links by saving the workbook to a new file.
// AI Prompts: Generate C# code that lists every linked OLE object in an Excel file and prints its original source path using Aspose.Cells. | Create a method that replaces the drive letter in ObjectSourceFullName of linked OLE objects from C: to D: and saves the workbook. | Explain best practices for handling exceptions when updating OleObject.ObjectSourceFullName for linked objects.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Shows how to load a workbook, iterate worksheets, filter OLE objects using OleObject.IsLink, read and change the ObjectSourceFullName of linked OLE items, and save the updated file.
    public class OleObjectIsLinkDemo
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all OLE objects in the current worksheet
                    foreach (OleObject ole in sheet.OleObjects)
                    {
                        // Process only linked OLE objects
                        if (!ole.IsLink)
                        {
                            // Skip embedded OLE objects
                            Console.WriteLine("Skipping embedded OLE object.");
                            continue;
                        }

                        // Display the original linked file path
                        Console.WriteLine($"Linked OLE object found. Original path: {ole.ObjectSourceFullName}");

                        // Example modification: change drive letter from C: to D:
                        string updatedPath = ole.ObjectSourceFullName.Replace("C:", "D:");
                        ole.ObjectSourceFullName = updatedPath;

                        // Show the updated path
                        Console.WriteLine($"Updated linked path: {ole.ObjectSourceFullName}");
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved with updated linked OLE objects: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
