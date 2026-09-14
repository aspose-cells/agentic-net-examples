// Title: Find Excel worksheets that contain only shapes (no cell data) using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that returns a collection of worksheet names where Cells.MaxDataRow is -1 and the Shapes collection is not empty. | Create a function in a .NET console app that scans an Excel file and prints the names of sheets that have no cell entries but contain at least one drawing object using Aspose.Cells. | Generate a C# example that loads a workbook, checks each worksheet for the absence of data and the presence of shapes, and outputs the matching sheet titles.
// Common Searches: Aspose.Cells C# find worksheets that have only drawings and no data | How to list Excel sheets with MaxDataRow -1 and shapes count greater than zero using Aspose.Cells | Identify empty worksheets containing shapes in a .xlsx file with Aspose.Cells .NET | C# code example to detect sheets with no cell values but with shape objects in Aspose.Cells
// Tags: filter worksheets by MaxDataRow Aspose.Cells | detect shape-only sheets Aspose.Cells | worksheet shapes count check C# | list empty worksheets with drawings Aspose.Cells | identify worksheets without cell data Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Loads a workbook, iterates through each worksheet, and collects the names of sheets where MaxDataRow equals -1 and the Shapes collection has at least one item, then prints those worksheet names.
class IdentifyShapeOnlyWorksheets
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // List to hold names of worksheets that contain only shapes
            List<string> shapeOnlySheets = new List<string>();

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // MaxDataRow returns -1 when there is no data in the sheet
                bool hasNoData = sheet.Cells.MaxDataRow == -1;

                // ShapeCollection.Count gives the number of shapes on the sheet
                bool hasShapes = sheet.Shapes.Count > 0;

                // Identify sheets that have no data but contain at least one shape
                if (hasNoData && hasShapes)
                {
                    shapeOnlySheets.Add(sheet.Name);
                }
            }

            // Output the result
            Console.WriteLine("Worksheets containing only shapes:");
            foreach (string name in shapeOnlySheets)
            {
                Console.WriteLine("- " + name);
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
