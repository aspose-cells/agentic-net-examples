// Title: Batch Insert Linked Pictures into Excel Cells from a CSV using Aspose.Cells for .NET
// Description: Reads a CSV where each line contains a cell reference and an image URL, validates the data, creates a workbook, and uses Shapes.AddLinkedPicture to place a 100 × 100 px linked picture at the specified cell. The workbook is then saved as an XLSX file.
// Keywords: Aspose.Cells | C# | .NET | AddLinkedPicture | linked picture | batch insert images | CSV to Excel | Excel automation | worksheet shapes | bulk image import | Excel file generation
// Common Searches: Aspose.Cells batch insert pictures from CSV | Add linked images to Excel cells C# | Read CSV and place pictures in Excel with Aspose | Bulk load images into worksheet using Shapes.AddLinkedPicture | C# example for inserting multiple pictures into Excel
// Developer Intent: Insert multiple linked pictures into specific Excel cells based on a CSV list of cell references and image URLs.
// Use Cases: Create a product catalog where each SKU cell is paired with its image URL from a CSV. | Generate a marketing report that embeds web‑hosted chart images into designated cells. | Automate a dashboard that pulls promotional images from a data feed and places them in predefined locations.
// AI Prompts: Write a C# method that reads a CSV of cell references and image URLs and uses Aspose.Cells to add linked pictures with custom dimensions. | Extend the example to support variable picture sizes defined by an extra column in the CSV. | Add robust error handling that logs invalid URLs, skips malformed lines, and outputs a processing summary.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchInsertPictures
{
    // Reads a CSV where each line contains a cell reference and an image URL, validates the data, creates a workbook, and uses Shapes.AddLinkedPicture to place a 100 × 100 px linked picture at the specified cell. The workbook is then saved as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the CSV file containing cell references and image URLs.
            // Expected format per line: CellReference,ImageUrl
            // Example: B2,https://example.com/image1.jpg
            string csvPath = "images.csv";

            try
            {
                // Verify that the CSV file exists before attempting to read it.
                if (!File.Exists(csvPath))
                {
                    Console.WriteLine($"CSV file not found: {Path.GetFullPath(csvPath)}");
                    return;
                }

                // Create a new workbook (lifecycle create rule)
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Read and process each line of the CSV file
                foreach (string line in File.ReadLines(csvPath))
                {
                    // Skip empty lines
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    // Split the line into cell reference and image URL
                    string[] parts = line.Split(new[] { ',' }, 2);
                    if (parts.Length != 2)
                        continue; // Invalid line format, ignore

                    string cellRef = parts[0].Trim();
                    string imageUrl = parts[1].Trim();

                    // Obtain the cell to determine its row and column indices (zero‑based)
                    Cell cell;
                    try
                    {
                        cell = worksheet.Cells[cellRef];
                    }
                    catch
                    {
                        // Invalid cell reference, ignore this line
                        continue;
                    }

                    int topRow = cell.Row;
                    int leftColumn = cell.Column;

                    // Define picture size in pixels (adjust as needed)
                    int pictureHeight = 100;
                    int pictureWidth = 100;

                    // Add a linked picture to the worksheet at the specified cell location
                    // Using ShapeCollection.AddLinkedPicture method (provided rule)
                    worksheet.Shapes.AddLinkedPicture(topRow, leftColumn, pictureHeight, pictureWidth, imageUrl);
                }

                // Save the workbook (lifecycle save rule)
                string outputPath = "output_with_pictures.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
