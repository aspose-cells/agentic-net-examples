// Title: C# – Update the display label of a linked OLE object in Excel using Aspose.Cells
// Description: Demonstrates how to add a linked OLE object (e.g., a Word document) to a worksheet, assign a custom display label via the OleObject.Label property, save the workbook, reload it, and confirm that the label persists—all with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# OLE object label | set OLE object display text .NET | linked OLE object Excel Aspose | OleObject.Label property | programmatically change OLE label
// Common Searches: how to set label for linked OLE object in Excel using Aspose.Cells | Aspose.Cells change OLE object display name C# | read OLE object label after saving workbook Aspose | C# update OleObject.Label property | add Word document OLE link with custom label Aspose.Cells
// Developer Intent: Assign or modify the visible label of a linked OLE object in an Excel workbook through Aspose.Cells for .NET.
// Use Cases: Insert a linked Word file into a sheet and give it a meaningful caption for end‑users. | Batch‑process existing workbooks to rename OLE object labels based on file names or business rules. | Validate that a custom label survives workbook serialization and can be read back programmatically.
// AI Prompts: Generate C# code that sets the Label property of an OleObject after adding it with Aspose.Cells. | Show how to iterate through all OleObjects in a workbook and assign each a label derived from its source path. | Explain how to verify the label of a linked OLE object after loading a saved workbook with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectLabelUpdateDemo
{
    // Demonstrates how to add a linked OLE object (e.g., a Word document) to a worksheet, assign a custom display label via the OleObject.Label property, save the workbook, reload it, and confirm that the label persists—all with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to a source file that will be linked as an OLE object.
                string sourceFilePath = @"C:\Temp\SampleDocument.docx";

                if (!File.Exists(sourceFilePath))
                {
                    Console.WriteLine($"Source file not found: {sourceFilePath}");
                    return;
                }

                // -----------------------------------------------------------------
                // Create a new workbook and add a linked OLE object with a label.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Use null for the placeholder image; Aspose.Cells will apply a default icon.
                byte[] placeholderImage = null;

                // Add the OLE object at cell B2 (row 1, column 1) with the placeholder image.
                int oleIndex = sheet.OleObjects.Add(1, 1, 200, 200, placeholderImage);
                OleObject ole = sheet.OleObjects[oleIndex];

                // Configure the OLE object to link to the external file.
                ole.IsLink = true;
                ole.ObjectSourceFullName = sourceFilePath;
                ole.ProgID = "Word.Document.12"; // ProgID for .docx

                // Set the display label that will appear on the worksheet.
                ole.Label = "Project Specification Document";

                // Ensure the output directory exists.
                string outputPath = @"C:\Temp\OleObjectLabelDemo.xlsx";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to disk.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");

                // -----------------------------------------------------------------
                // Load the saved workbook and verify the label of the OLE object.
                // -----------------------------------------------------------------
                if (!File.Exists(outputPath))
                {
                    Console.WriteLine($"Saved workbook not found: {outputPath}");
                    return;
                }

                Workbook loadedWorkbook = new Workbook(outputPath);
                OleObject loadedOle = loadedWorkbook.Worksheets[0].OleObjects[0];

                Console.WriteLine("Loaded OLE Object Label: " + loadedOle.Label);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
