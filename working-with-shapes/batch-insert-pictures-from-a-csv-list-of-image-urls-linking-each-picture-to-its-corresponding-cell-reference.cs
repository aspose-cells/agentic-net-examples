// Title: How to batch insert pictures from a CSV of image URLs into specific Excel cells using Aspose.Cells for .NET (C#)
// Common Searches: C# read CSV of image URLs and add each picture to a designated cell with Aspose.Cells | Insert multiple pictures into an Excel worksheet from URLs listed in a CSV file using Aspose.Cells .NET | Set picture placement to MoveAndSize when adding images to cells with Aspose.Cells C# | Convert Excel cell reference like B3 to row and column indexes for picture insertion in Aspose.Cells | Download images with HttpClient and embed them in an Aspose.Cells workbook
// Tags: batch picture insertion from CSV Aspose.Cells | add picture to worksheet from MemoryStream Aspose.Cells | picture placement MoveAndSize Aspose.Cells | cell reference to row column conversion C# | download image via HttpClient for Excel insertion

using System;
using System.IO;
using System.Net.Http;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example reads a CSV file containing cell references and image URLs, downloads each image with HttpClient into a MemoryStream, converts the cell reference to zero‑based row and column indexes, inserts the picture into the specified cell using Aspose.Cells, sets its placement to MoveAndSize so it moves and resizes with the cell, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the CSV file containing cell reference and image URL (e.g., A1,https://example.com/img.png)
            string csvPath = "images.csv";

            // Verify that the CSV file exists to avoid FileNotFoundException
            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"CSV file not found: {Path.GetFullPath(csvPath)}");
                return;
            }

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // HttpClient for downloading images
            using (HttpClient client = new HttpClient())
            {
                foreach (string line in File.ReadAllLines(csvPath))
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    // Expected CSV format: CellReference,ImageUrl
                    string[] parts = line.Split(new[] { ',' }, 2);
                    if (parts.Length != 2)
                        continue;

                    string cellRef = parts[0].Trim();
                    string imageUrl = parts[1].Trim();

                    // Convert cell reference (e.g., "B3") to zero‑based row and column indexes
                    CellReferenceToIndexes(cellRef, out int row, out int column);

                    try
                    {
                        // Download the image into a memory stream
                        byte[] imageBytes = client.GetByteArrayAsync(imageUrl).Result;
                        using (MemoryStream ms = new MemoryStream(imageBytes))
                        {
                            // Insert the picture at the specified cell
                            int pictureIndex = sheet.Pictures.Add(row, column, ms);

                            // Link the picture to the cell so it moves and resizes with the cell
                            sheet.Pictures[pictureIndex].Placement = PlacementType.MoveAndSize;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to process image for cell {cellRef}: {ex.Message}");
                    }
                }
            }

            // Save the workbook to a file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    // Helper: converts an Excel cell reference (e.g., "C5") to zero‑based row and column indexes
    static void CellReferenceToIndexes(string cellRef, out int row, out int column)
    {
        int i = 0;
        while (i < cellRef.Length && Char.IsLetter(cellRef[i]))
            i++;

        string colPart = cellRef.Substring(0, i).ToUpper();
        string rowPart = cellRef.Substring(i);

        column = 0;
        foreach (char c in colPart)
        {
            column = column * 26 + (c - 'A' + 1);
        }
        column--; // zero‑based column index

        row = int.Parse(rowPart) - 1; // zero‑based row index
    }
}
