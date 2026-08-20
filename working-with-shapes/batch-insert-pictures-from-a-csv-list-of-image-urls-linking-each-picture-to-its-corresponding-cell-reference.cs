// Title: Batch Insert Linked Pictures from CSV URLs into Excel Cells with Aspose.Cells for .NET
// Description: Creates a workbook, imports a CSV where column A contains cell addresses and column B holds image URLs, converts each address to row/column indices, and uses Shapes.AddLinkedPicture to place a 100 × 100 px linked image in the target cell. The workbook is then saved with all pictures embedded.
// Keywords: Aspose.Cells | .NET | linked picture | CSV import | batch image insertion | Excel cell address | Shapes.AddLinkedPicture | image URL | automated Excel graphics | worksheet picture placement
// Common Searches: Aspose.Cells add linked picture from URL | batch insert images into Excel using .NET | place picture in cell address from CSV | how to import image URLs into Excel worksheet | automate picture placement with Aspose.Cells
// Developer Intent: Programmatically add multiple linked images to specific Excel cells based on a CSV mapping of cell references to image URLs.
// Use Cases: Generate a product catalog where each SKU cell automatically shows its online image. | Create a regional sales dashboard that inserts flag icons into header cells using a URL list. | Build a marketing report that populates predefined cells with brand logos from a CSV configuration.
// AI Prompts: Write a C# method that reads a CSV of cell addresses and image URLs and inserts linked pictures into the corresponding cells using Aspose.Cells. | Modify the sample to accept width and height columns in the CSV and set each picture's dimensions accordingly. | Add error handling that logs invalid URLs, skips failed rows, and continues processing the remaining entries.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, imports a CSV where column A contains cell addresses and column B holds image URLs, converts each address to row/column indices, and uses Shapes.AddLinkedPicture to place a 100 × 100 px linked image in the target cell. The workbook is then saved with all pictures embedded.
class BatchInsertPictures
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the CSV file containing cell addresses and image URLs
            string csvPath = "images.csv";

            // Import CSV only if the file exists
            if (File.Exists(csvPath))
            {
                // Import CSV: Column A = cell address, Column B = image URL
                worksheet.Cells.ImportCSV(csvPath, ",", true, 0, 0);
            }
            else
            {
                Console.WriteLine($"CSV file not found: {csvPath}");
                // Continue with an empty worksheet or exit as needed
                // Here we simply proceed without inserting pictures
                workbook.Save("output_with_pictures.xlsx");
                return;
            }

            // Determine the last row that contains data
            int lastRow = worksheet.Cells.MaxDataRow;

            // Iterate through each row of the CSV data
            for (int row = 0; row <= lastRow; row++)
            {
                // Read cell address and image URL from the imported CSV
                string cellAddress = worksheet.Cells[row, 0].StringValue?.Trim();
                string imageUrl = worksheet.Cells[row, 1].StringValue?.Trim();

                // Skip rows with missing data
                if (string.IsNullOrEmpty(cellAddress) || string.IsNullOrEmpty(imageUrl))
                    continue;

                // Convert the cell address (e.g., "C5") to row and column indices
                Cell targetCell = worksheet.Cells[cellAddress];
                int targetRow = targetCell.Row;
                int targetColumn = targetCell.Column;

                // Add a linked picture at the specified cell (size: 100x100 pixels)
                worksheet.Shapes.AddLinkedPicture(targetRow, targetColumn, 100, 100, imageUrl);
            }

            // Save the workbook with the inserted pictures
            workbook.Save("output_with_pictures.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
