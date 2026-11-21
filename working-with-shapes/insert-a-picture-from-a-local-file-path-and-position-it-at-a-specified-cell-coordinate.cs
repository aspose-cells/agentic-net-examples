// Title: Insert a local JPEG image into a specific Excel cell using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells add picture to cell B2 from file path | How to anchor an image to a worksheet cell with Aspose.Cells .NET | Insert local image into Excel worksheet using Aspose.Cells Pictures.Add | Check file existence before adding picture in Aspose.Cells C# | Convert cell address to row and column indices Aspose.Cells
// Tags: pictures.add method for inserting images | jpeg image insertion into Excel cell .NET | image anchoring to specific cell Aspose.Cells | file existence check prior to picture add C# | cell address to row column conversion Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, verifies a local JPEG file, converts a target cell address (e.g., B2) to row and column indices, inserts the picture at that cell using sheet.Pictures.Add, and saves the workbook, handling any errors that may occur.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Local picture file path
            string picturePath = @"C:\Images\sample.jpg";

            // Verify that the picture file exists before adding
            if (File.Exists(picturePath))
            {
                // Target cell where the picture will be anchored (e.g., B2)
                string targetCell = "B2";

                // Convert cell name to row and column indices
                int row = sheet.Cells[targetCell].Row;
                int column = sheet.Cells[targetCell].Column;

                // Insert the picture and anchor it to the specified cell
                sheet.Pictures.Add(row, column, picturePath);
            }
            else
            {
                Console.WriteLine($"Picture file not found: {picturePath}. Skipping picture insertion.");
            }

            // Save the workbook
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
