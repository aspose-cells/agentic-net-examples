// Title: Batch set a standard column width for every worksheet in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, iterates all worksheets, sets each sheet's Cells.StandardWidth to a specified value, and saves the workbook. | Show how to apply a uniform default column width to every worksheet in a workbook programmatically with Aspose.Cells for .NET. | Create a console application that reads a workbook, assigns a consistent column width (e.g., 18.25 characters) to all sheets, and writes the updated file.
// Common Searches: how to set default column width for all sheets using Aspose.Cells C# | batch update column width across multiple worksheets Aspose.Cells .NET example | C# loop through workbook worksheets and apply StandardWidth property
// Tags: set StandardWidth across worksheets | batch column width adjustment Aspose.Cells | uniform column sizing Excel .NET | apply default column width workbook | iterate worksheets Aspose.Cells C#

using System;
using Aspose.Cells;

namespace AsposeCellsBatchStandardWidth
{
    // // Loads an existing Excel workbook, sets the default column width to 18.25 characters for every worksheet via the Cells.StandardWidth property, and saves the modified file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook (replace with actual file path)
            string inputPath = "input.xlsx";

            // Path to the destination workbook
            string outputPath = "output.xlsx";

            // Create a Workbook object and load the existing file
            Workbook workbook = new Workbook(inputPath);

            // Desired standard column width (in characters)
            double standardWidth = 18.25;

            // Apply the standard width to every worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Set the default column width for the current worksheet
                sheet.Cells.StandardWidth = standardWidth;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
    }
}
