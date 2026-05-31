using System;
using System.IO;
using System.Net.Http;
using Aspose.Cells;

namespace BatchInsertPicturesApp
{
    class BatchInsertPictures
    {
        static void Main()
        {
            try
            {
                // Path to the CSV file containing cell references and image URLs.
                string csvPath = "images.csv";

                // Verify that the CSV file exists before attempting to read it.
                if (!File.Exists(csvPath))
                {
                    Console.WriteLine($"CSV file not found: {csvPath}");
                    return;
                }

                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Prepare an HttpClient for downloading images.
                using (HttpClient httpClient = new HttpClient())
                {
                    // Read all non‑empty lines from the CSV file.
                    foreach (string line in File.ReadAllLines(csvPath))
                    {
                        if (string.IsNullOrWhiteSpace(line))
                            continue; // Skip empty lines.

                        // Split the line into cell reference and image URL.
                        string[] parts = line.Split(new[] { ',' }, 2);
                        if (parts.Length < 2)
                            continue; // Skip malformed lines.

                        string cellRef = parts[0].Trim();
                        string imageUrl = parts[1].Trim();

                        // Download the image data synchronously with error handling.
                        byte[] imageData;
                        try
                        {
                            imageData = httpClient.GetByteArrayAsync(imageUrl).Result;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to download image from '{imageUrl}': {ex.Message}");
                            continue;
                        }

                        // Create a memory stream from the downloaded bytes.
                        using (MemoryStream imageStream = new MemoryStream(imageData))
                        {
                            // Resolve the cell to obtain its zero‑based row and column indices.
                            Cell targetCell = worksheet.Cells[cellRef];
                            int row = targetCell.Row;
                            int column = targetCell.Column;

                            // Add the picture to the worksheet anchored at the target cell.
                            worksheet.Pictures.Add(row, column, imageStream);
                        }
                    }
                }

                // Save the workbook with all inserted pictures.
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}