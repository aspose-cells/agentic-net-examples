// Title: Flag Excel worksheets that contain only shapes (no data rows) using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates all worksheets, checks if Cells.MaxDataRow equals -1 and Shapes.Count > 0, then adds a custom property named "ShapeOnlyContent" set to "True" and writes "Shape‑only content detected" into cell A1. | Update an existing workbook using Aspose.Cells to mark sheets that have no data rows but contain shapes by inserting a flag property and a visible label, then save the modified file. | Create a reusable C# method that returns a list of worksheet names where MaxDataRow is -1 and the shape collection is not empty, using Aspose.Cells.
// Common Searches: asp.net detect worksheets that contain only shapes with Aspose.Cells | how to store a flag in an Excel sheet when it has shapes but no data rows in C# | Aspose.Cells check for empty data rows while shapes exist in a worksheet | flag shape‑only sheets in an Excel workbook using Aspose.Cells .NET | list worksheet names with MaxDataRow -1 and Shapes.Count > 0 Aspose.Cells
// Tags: shape‑only worksheet detection Aspose.Cells | shape‑only flag property Excel | MaxDataRow -1 and Shapes.Count validation | flag worksheets without data rows .NET | write label to cell A1 Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program loads an Excel workbook, scans each worksheet, and when a sheet has no data rows (Cells.MaxDataRow == -1) but contains shapes (Shapes.Count > 0), it adds or updates a custom property "ShapeOnlyContent" set to "True" and writes a visible note "Shape‑only content detected" into cell A1 before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // MaxDataRow == -1 means the sheet has no data rows
                // Shapes.Count > 0 means there are shapes present
                if (sheet.Cells.MaxDataRow == -1 && sheet.Shapes.Count > 0)
                {
                    // Add or update a custom property to flag the worksheet
                    var existingProp = sheet.CustomProperties["ShapeOnlyContent"];
                    if (existingProp != null)
                    {
                        // Store boolean as string because CustomProperty.Value expects a string
                        existingProp.Value = true.ToString();
                    }
                    else
                    {
                        // Add expects a string value; store boolean as string
                        sheet.CustomProperties.Add("ShapeOnlyContent", true.ToString());
                    }

                    // Write a note in cell A1 for visual indication
                    Cell flagCell = sheet.Cells["A1"];
                    flagCell.PutValue("Shape‑only content detected");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
