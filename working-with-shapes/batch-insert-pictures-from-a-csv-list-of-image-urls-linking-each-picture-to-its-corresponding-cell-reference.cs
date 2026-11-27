// Title: Insert multiple pictures into Excel cells from a CSV of image URLs and add hyperlinks using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells insert pictures from CSV list of URLs into specific cells | How to add hyperlinks to pictures inserted with Aspose.Cells in .NET | Batch download images and embed them in an Excel worksheet using Aspose.Cells | Parse A1 cell reference to row and column indexes with Aspose.Cells C# | Place pictures as free‑floating over cells using the Aspose.Cells API
// Tags: batch picture insertion Aspose.Cells | image URL download Aspose.Cells C# | hyperlink on inserted picture Aspose.Cells | freefloating picture placement Excel | A1 cell reference parsing Aspose.Cells | CSV driven image embedding Aspose.Cells

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace BatchInsertPicturesFromCsv
{
    // The example reads a CSV file containing cell references and image URLs, downloads each image, inserts it as a picture at the target cell in a new workbook, sets the picture to free‑floating placement, attaches a hyperlink to the original URL, and saves the result as an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        // Entry point – async to allow awaiting HTTP calls
        static async Task Main(string[] args)
        {
            try
            {
                // Path to the CSV file (first argument or default)
                string csvPath = args.Length > 0 ? args[0] : "images.csv";

                if (!File.Exists(csvPath))
                {
                    Console.WriteLine($"CSV file not found: {csvPath}");
                    return;
                }

                // Load CSV data (expected format: CellReference,ImageUrl)
                List<(string Cell, string Url)> entries = LoadCsv(csvPath);
                if (entries.Count == 0)
                {
                    Console.WriteLine("No entries found in CSV.");
                    return;
                }

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // HttpClient for downloading images
                using (HttpClient httpClient = new HttpClient())
                {
                    foreach (var (cellRef, imageUrl) in entries)
                    {
                        try
                        {
                            // Download image data
                            byte[] imageData = await httpClient.GetByteArrayAsync(imageUrl);
                            using (MemoryStream ms = new MemoryStream(imageData))
                            {
                                // Convert cell reference (e.g., "B2") to zero‑based row/column indices
                                if (!CellReferenceToIndexes(cellRef, out int row, out int col))
                                {
                                    Console.WriteLine($"Invalid cell reference '{cellRef}'. Skipping.");
                                    continue;
                                }

                                // Insert picture at the specified cell
                                int picIndex = sheet.Pictures.Add(row, col, ms);
                                Picture picture = sheet.Pictures[picIndex];

                                // Make the picture free‑floating so it stays over the cell
                                picture.Placement = PlacementType.FreeFloating;

                                // Add a hyperlink to the picture (Hyperlink property is read‑only, but its members are settable)
                                picture.Hyperlink.Address = imageUrl;
                                picture.Hyperlink.ScreenTip = $"Image link for {cellRef}";

                                Console.WriteLine($"Inserted picture for cell {cellRef} from {imageUrl}");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to process URL '{imageUrl}' for cell '{cellRef}': {ex.Message}");
                        }
                    }
                }

                // Save the workbook
                string outputPath = "output.xlsx";
                try
                {
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved to {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Simple CSV loader – assumes no commas inside fields and no header line
        private static List<(string Cell, string Url)> LoadCsv(string path)
        {
            var list = new List<(string Cell, string Url)>();
            foreach (var line in File.ReadAllLines(path))
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                string[] parts = line.Split(new[] { ',' }, 2);
                if (parts.Length != 2)
                {
                    Console.WriteLine($"Skipping malformed line: {line}");
                    continue;
                }

                string cell = parts[0].Trim();
                string url = parts[1].Trim();
                if (!string.IsNullOrEmpty(cell) && !string.IsNullOrEmpty(url))
                {
                    list.Add((cell, url));
                }
            }
            return list;
        }

        // Convert an A1 style cell reference to zero‑based row and column indexes
        private static bool CellReferenceToIndexes(string cellRef, out int row, out int column)
        {
            row = column = -1;
            if (string.IsNullOrEmpty(cellRef))
                return false;

            // Separate letters and numbers
            int i = 0;
            while (i < cellRef.Length && char.IsLetter(cellRef[i]))
                i++;

            if (i == 0 || i == cellRef.Length)
                return false;

            string colPart = cellRef.Substring(0, i).ToUpperInvariant();
            string rowPart = cellRef.Substring(i);

            // Convert column letters to number (A=0, B=1, ..., Z=25, AA=26, etc.)
            int colNum = 0;
            foreach (char ch in colPart)
            {
                if (ch < 'A' || ch > 'Z')
                    return false;
                colNum = colNum * 26 + (ch - 'A' + 1);
            }
            column = colNum - 1; // zero‑based

            // Convert row part to integer
            if (!int.TryParse(rowPart, out int rowNum) || rowNum <= 0)
                return false;
            row = rowNum - 1; // zero‑based

            return true;
        }
    }
}
