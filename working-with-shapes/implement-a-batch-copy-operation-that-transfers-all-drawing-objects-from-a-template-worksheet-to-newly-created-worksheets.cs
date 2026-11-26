// Title: Batch copy all drawing objects from a template worksheet to multiple new worksheets with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy drawings from a template sheet to several new worksheets | How to duplicate all shapes when adding copies of a worksheet in Aspose.Cells .NET | Clear cell contents but keep embedded drawings using Aspose.Cells for C# | Create multiple worksheets from a template and preserve its drawings with Aspose.Cells | Remove template sheet after copying drawings in Aspose.Cells .NET
// Tags: copy worksheet drawings Aspose.Cells .NET | duplicate shapes across multiple sheets C# | clear cells preserve drawings Aspose.Cells | add copy of template worksheet Aspose.Cells | batch copy drawing objects .NET | remove template sheet after copying Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace BatchCopyDrawings
{
    // The example loads a template workbook, adds copies of its first worksheet for each target name, clears only the cell data while retaining all drawing objects, optionally deletes the original template sheet, and saves the result as a new XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths – adjust as needed. The template workbook must exist.
            string templatePath = "Template.xlsx";
            string outputPath = "Result.xlsx";

            // Verify that the template file exists.
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Error: Template file \"{templatePath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook that contains the template worksheet.
                Workbook workbook = new Workbook(templatePath);

                // Assume the template worksheet is the first sheet (index 0).
                int templateIndex = 0;
                Worksheet templateSheet = workbook.Worksheets[templateIndex];
                Console.WriteLine($"Template worksheet \"{templateSheet.Name}\" loaded successfully.");

                // Define the names of the new worksheets that will receive the drawings.
                string[] newSheetNames = new string[]
                {
                    "Copy_1",
                    "Copy_2",
                    "Copy_3"
                };

                foreach (string newName in newSheetNames)
                {
                    // Add a copy of the template worksheet. This copies all drawings, formats, etc.
                    int newIndex = workbook.Worksheets.AddCopy(templateIndex);
                    Worksheet newSheet = workbook.Worksheets[newIndex];

                    // Rename the newly created worksheet.
                    newSheet.Name = newName;

                    // Clear all cell contents while preserving drawings.
                    // Cells.Clear clears only the data, not the drawing objects.
                    newSheet.Cells.Clear();

                    Console.WriteLine($"Created worksheet \"{newName}\" with drawings copied from template.");
                }

                // Optionally, remove the original template sheet if it should not be part of the final file.
                // Uncomment the following line to delete the template sheet:
                // workbook.Worksheets.RemoveAt(templateIndex);

                // Save the resulting workbook.
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An exception occurred during processing:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
