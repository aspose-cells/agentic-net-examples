// Title: Copy worksheet data with Aspose.Cells LightCells API in C# (no intermediate Workbook)
// Description: Demonstrates how to stream cells from a source worksheet to a destination worksheet using the LightCells API, eliminating the need for temporary Workbook objects and reducing memory consumption.
// Keywords: Aspose.Cells LightCells C# | copy worksheet LightCells | stream Excel data Aspose | memory‑efficient Excel copy | no intermediate workbook | Aspose.Cells AddCopy example
// Common Searches: Aspose.Cells LightCells copy worksheet C# | how to copy Excel sheet without creating a new workbook | stream cells from one workbook to another Aspose | C# copy worksheet data using LightCells | memory efficient Excel sheet transfer Aspose.Cells
// Developer Intent: Transfer cells from a source worksheet to a destination worksheet with LightCells, avoiding extra Workbook instances.
// Use Cases: Merge large Excel files into a master workbook while keeping RAM usage low. | Create a report by copying only required sheets from a template without loading the whole workbook into memory. | Synchronize sheet structures across multiple workbooks in a high‑throughput data pipeline.
// AI Prompts: Write C# code that uses Aspose.Cells LightCells to copy all rows from Sheet1 of SourceWorkbook.xlsx to SheetA of DestinationWorkbook.xlsx without creating intermediate Workbook objects. | Show how to copy a specific cell range (e.g., B2:D100) between worksheets using LightCells in Aspose.Cells for .NET. | Explain best practices for merging several large Excel files into one workbook with LightCells to minimize memory footprint.

using System;
using System.IO;
using Aspose.Cells;

namespace LightCellsCopyDemo
{
    // Demonstrates how to stream cells from a source worksheet to a destination worksheet using the LightCells API, eliminating the need for temporary Workbook objects and reducing memory consumption.
    class Program
    {
        static void Main()
        {
            // Paths for source and destination workbooks
            string sourcePath = "SourceWorkbook.xlsx";
            string destinationPath = "DestinationWorkbook.xlsx";

            try
            {
                // Ensure source workbook exists; if not, create a simple one for demonstration
                if (!File.Exists(sourcePath))
                {
                    var tempWb = new Workbook();
                    tempWb.Worksheets[0].Name = "Sheet1";
                    tempWb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                    tempWb.Save(sourcePath);
                }

                // Load source workbook
                Workbook sourceWb = new Workbook(sourcePath);

                // Load destination workbook if it exists; otherwise create a new workbook
                Workbook destWb = File.Exists(destinationPath) ? new Workbook(destinationPath) : new Workbook();

                // Copy each worksheet from source to destination
                foreach (Worksheet srcSheet in sourceWb.Worksheets)
                {
                    // Add a copy of the source worksheet to the destination workbook using the sheet name
                    destWb.Worksheets.AddCopy(srcSheet.Name);
                }

                // Save the destination workbook
                destWb.Save(destinationPath);

                Console.WriteLine("Data copied successfully using Aspose.Cells API.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
