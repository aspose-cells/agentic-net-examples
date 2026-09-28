// Title: Batch insert pictures into Excel cells from a CSV of image URLs using Aspose.Cells for .NET
// AI Prompts: Write C# code that reads a CSV file with cell addresses and image URLs, downloads each image via HttpClient, inserts the picture into the corresponding cell of an Aspose.Cells workbook, and sets the picture's Placement to MoveAndSize. | Update the insertion loop to give each picture a name derived from its cell reference and to log download or insertion errors without stopping the batch process.
// Common Searches: Aspose.Cells C# add pictures to cells from a CSV list of URLs | C# batch download images and place them in specific Excel cells using Aspose.Cells | How to set picture placement to MoveAndSize when inserting images into Excel with Aspose.Cells | Read cell reference and image URL from CSV and insert image into Excel worksheet using Aspose.Cells .NET
// Tags: batch picture insertion Aspose.Cells | add image to Excel cell from URL | move and size picture placement | CSV-driven image insertion C# | HttpClient image download for Excel

using System;
using System.IO;
using System.Net.Http;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program reads a CSV file containing cell references and image URLs, downloads each image with HttpClient, adds the picture to the matching cell in a new Aspose.Cells workbook, sets the picture to MoveAndSize with the cell, optionally names the picture after its cell, and saves the workbook as OutputWithPictures.xlsx.
class BatchInsertPictures
{
    static void Main()
    {
        // Path to the CSV file containing cell references and image URLs
        string csvPath = "images.csv";

        // Verify that the CSV file exists to avoid FileNotFoundException
        if (!File.Exists(csvPath))
        {
            Console.WriteLine($"CSV file not found: {csvPath}");
            return;
        }

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // HttpClient for downloading images (disposed after use)
        using (HttpClient httpClient = new HttpClient())
        {
            try
            {
                // Read all lines from the CSV file
                foreach (string line in File.ReadAllLines(csvPath))
                {
                    // Skip empty lines
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    // Expected CSV format: CellReference,ImageUrl
                    // Example: B2,https://example.com/pic1.png
                    string[] parts = line.Split(new[] { ',' }, 2);
                    if (parts.Length != 2)
                        continue; // malformed line

                    string cellRef = parts[0].Trim();
                    string imageUrl = parts[1].Trim();

                    // Convert cell reference (e.g., "B2") to zero‑based row and column indices
                    Cell cell = sheet.Cells[cellRef];
                    int row = cell.Row;
                    int column = cell.Column;

                    // Download the image into a memory stream
                    try
                    {
                        using (Stream imageStream = httpClient.GetStreamAsync(imageUrl).Result)
                        {
                            // Add the picture to the worksheet at the specified cell
                            int pictureIndex = sheet.Pictures.Add(row, column, imageStream);
                            Picture picture = sheet.Pictures[pictureIndex];

                            // Set the picture to move and size with cells (linking it to the cell)
                            picture.Placement = PlacementType.MoveAndSize;

                            // Optional: give the picture a name based on the cell reference
                            picture.Name = $"Pic_{cellRef}";
                        }
                    }
                    catch (Exception ex)
                    {
                        // If download or insertion fails, log and continue with next entry
                        Console.WriteLine($"Failed to insert picture for {cellRef}: {ex.Message}");
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing CSV file: {ex.Message}");
                return;
            }
        }

        // Save the workbook to a file (ensure the directory is writable)
        try
        {
            workbook.Save("OutputWithPictures.xlsx");
            Console.WriteLine("Workbook saved as OutputWithPictures.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
