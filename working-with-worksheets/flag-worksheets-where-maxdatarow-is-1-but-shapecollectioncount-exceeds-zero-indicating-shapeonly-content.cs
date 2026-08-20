// Title: Flag Shape‑Only Worksheets (No Cell Data) with a Custom Property using Aspose.Cells for .NET
// Description: Load an Excel workbook, iterate each worksheet, detect sheets where Cells.MaxDataRow = -1 and Shapes.Count > 0, add a custom property "ShapeOnly" = true, and save the updated file.
// Keywords: Aspose.Cells | C# Excel | shape only worksheet | MaxDataRow | Shapes.Count | custom property | flag worksheet | drawing objects | worksheet metadata | .NET Excel automation
// Common Searches: Aspose.Cells detect worksheets with only drawings | C# flag Excel sheets that have no data but contain shapes | Add custom property to Excel sheet using Aspose.Cells | Check MaxDataRow and Shapes.Count in .NET | Identify shape‑only worksheets programmatically
// Developer Intent: Detect worksheets lacking cell data but containing drawings and mark them with a custom property.
// Use Cases: Automated content classification before publishing workbooks | Generate audit reports of drawing‑only sheets | Skip or treat shape‑only worksheets differently in downstream processes | Ensure compliance by tagging non‑data sheets | Integrate with document management systems to flag shape‑only content
// AI Prompts: Generate C# code with Aspose.Cells that scans each worksheet, checks Cells.MaxDataRow == -1 and Shapes.Count > 0, adds a custom property "ShapeOnly" = "true", and saves the file. | Explain how to combine MaxDataRow and Shapes.Count to identify shape‑only worksheets and use CustomProperties to flag them in Aspose.Cells for .NET. | Provide error‑handling and performance recommendations for flagging shape‑only worksheets in large Excel workbooks using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsShapeOnlyFlag
{
    // Load an Excel workbook, iterate each worksheet, detect sheets where Cells.MaxDataRow = -1 and Shapes.Count > 0, add a custom property "ShapeOnly" = true, and save the updated file.
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
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // MaxDataRow returns -1 when the worksheet contains no cell data
                    bool hasNoCellData = sheet.Cells.MaxDataRow == -1;

                    // ShapeCollection.Count gives the number of drawing objects on the sheet
                    bool hasShapes = sheet.Shapes.Count > 0;

                    // Flag the worksheet if it has shapes but no cell data
                    if (hasNoCellData && hasShapes)
                    {
                        // Add a custom property named "ShapeOnly" with value "true"
                        sheet.CustomProperties.Add("ShapeOnly", true.ToString());

                        // Write the sheet name to console for verification
                        Console.WriteLine($"Worksheet \"{sheet.Name}\" flagged as shape‑only content.");
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
