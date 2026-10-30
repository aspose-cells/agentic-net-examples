// Title: Batch copy all drawing objects from a template worksheet to multiple new worksheets using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells copy all shapes from one sheet to several new sheets | How to duplicate drawing objects across multiple worksheets with Aspose.Cells | Batch transfer of worksheet shapes using Aspose.Cells AddCopy method | Programmatically clone template sheet drawings to new worksheets in .NET | Copy shapes from a template workbook to multiple worksheets using Aspose.Cells for C#
// Tags: Aspose.Cells copy shapes between worksheets | AddCopy shape method C# | batch copy worksheet drawings .NET | duplicate template sheet drawing objects | clone worksheet shapes Aspose.Cells | transfer drawing objects to new worksheets

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a template workbook, creates a new workbook, adds several worksheets, and uses the Shape.AddCopy method to duplicate every drawing object from the template sheet onto each new sheet while preserving positions, then saves the result.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Verify template file exists
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file \"{templatePath}\" not found.");
                return;
            }

            // Load the template workbook
            Workbook templateWorkbook = new Workbook(templatePath);
            Worksheet templateWorksheet = templateWorkbook.Worksheets[0]; // assume first sheet is the template

            // Names of the new worksheets that will receive the copied drawing objects
            List<string> newWorksheetNames = new List<string> { "CopySheet1", "CopySheet2", "CopySheet3" };

            // Create a new workbook to hold the copied worksheets
            Workbook resultWorkbook = new Workbook();
            // Remove the default sheet created by Aspose.Cells
            resultWorkbook.Worksheets.Clear();

            // Iterate over each target worksheet name
            foreach (string sheetName in newWorksheetNames)
            {
                // Add a new worksheet with the specified name
                Worksheet targetWorksheet = resultWorkbook.Worksheets.Add(sheetName);

                // Copy each shape (drawing object) from the template worksheet to the new worksheet
                foreach (Shape shape in templateWorksheet.Shapes)
                {
                    // Preserve the original shape's position while copying
                    targetWorksheet.Shapes.AddCopy(
                        shape,
                        shape.UpperLeftRow,
                        shape.UpperLeftColumn,
                        shape.LowerRightRow,
                        shape.LowerRightColumn);
                }
            }

            // Ensure the directory for the result file exists
            string resultDir = Path.GetDirectoryName(Path.GetFullPath(resultPath));
            if (!string.IsNullOrEmpty(resultDir) && !Directory.Exists(resultDir))
            {
                Directory.CreateDirectory(resultDir);
            }

            // Save the resulting workbook
            resultWorkbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to \"{resultPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
