// Title: Detect and flag worksheets that contain only shapes (no cell data) using Aspose.Cells for .NET (C#)
// Description: This C# sample loads a workbook (or creates a new one if the file is missing), iterates through each worksheet, and when Cells.MaxDataRow equals -1 while Shapes.Count is greater than zero, it adds a string custom property called "ShapeOnlyContent" set to "true" and renames the sheet by appending "_ShapeOnly". The workbook is then saved with the changes.
// Keywords: Aspose.Cells | .NET | C# | worksheet shape detection | MaxDataRow | Shapes.Count | custom property | rename worksheet | shape-only sheet | Excel automation
// Common Searches: Aspose.Cells detect worksheets with only drawings | C# check if worksheet has no data but has shapes | Add custom property to Excel sheet using Aspose.Cells | Rename Excel worksheet based on shape content .NET | MaxDataRow -1 Aspose.Cells example
// Developer Intent: Identify worksheets that contain drawings but no cell values and programmatically mark them.
// Use Cases: Skip drawing‑only sheets when converting a workbook to PDF. | Generate a list of shape‑only worksheets for manual review. | Automate workbook cleanup by flagging and renaming sheets that contain only shapes. | Trigger downstream workflows based on a custom property flag.
// AI Prompts: Write C# code with Aspose.Cells to flag worksheets where MaxDataRow = -1 and Shapes.Count > 0, add a custom property "ShapeOnlyContent" = "true" and rename the sheet with "_ShapeOnly". | Show how to log the names of flagged worksheets to a console or text file in the same Aspose.Cells example. | Suggest an alternative approach using worksheet comments instead of custom properties to indicate shape‑only content. | Explain how to modify the sample to process all workbooks in a folder and produce a summary report of shape‑only sheets.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# sample loads a workbook (or creates a new one if the file is missing), iterates through each worksheet, and when Cells.MaxDataRow equals -1 while Shapes.Count is greater than zero, it adds a string custom property called "ShapeOnlyContent" set to "true" and renames the sheet by appending "_ShapeOnly". The workbook is then saved with the changes.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load workbook if the input file exists; otherwise create a new workbook
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file '{inputPath}' not found. A new workbook will be created.");
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // MaxDataRow == -1 indicates no cell data; Shapes.Count > 0 indicates drawing objects
                if (sheet.Cells.MaxDataRow == -1 && sheet.Shapes.Count > 0)
                {
                    // Add a custom property (value stored as string) and rename the sheet
                    sheet.CustomProperties.Add("ShapeOnlyContent", "true");
                    sheet.Name = sheet.Name + "_ShapeOnly";
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
