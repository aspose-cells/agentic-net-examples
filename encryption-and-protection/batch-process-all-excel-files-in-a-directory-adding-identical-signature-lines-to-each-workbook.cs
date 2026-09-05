// Title: Batch add a bold signature row to every Excel workbook in a folder using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to enumerate all .xls, .xlsx, and .xlsm files in a specified input directory, insert a bold "Signature:" label and a name on a new row after the last used row of the first worksheet, and save each modified workbook to a designated output directory. | Generate a C# routine with Aspose.Cells that loads each workbook from a folder, determines the maximum data row, adds a formatted signature line (bold label and plain value) on the next empty row, and writes the updated file to another folder while handling missing files and exceptions.
// Common Searches: aspocells c# add signature line to many Excel files | how to programmatically append a footer to each worksheet using Aspose.Cells | C# loop through .xlsx files and write bold label after last data row | automate adding author name to all Excel workbooks with Aspose.Cells | process all Excel files in a folder and add custom row using Aspose.Cells
// Tags: bulk add signature line Aspose.Cells | append custom row to each workbook C# | detect last used row Aspose.Cells worksheet | save modified workbook to output folder .NET | apply bold style to cell Aspose.Cells

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

// // This program scans a source folder for .xls, .xlsx, and .xlsm files, loads each workbook with Aspose.Cells, finds the last populated row in the first worksheet, inserts a bold "Signature:" label and a name on the next row, and saves the updated workbook to a target directory.
class BatchSignatureAdder
{
    // Signature text to be added to each workbook
    private const string SignatureLabel = "Signature:";
    private const string SignatureValue = "Your Name";

    static void Main()
    {
        // Directory containing the source Excel files
        string sourceDirectory = @"C:\InputExcel";

        // Directory where the modified workbooks will be saved
        string destinationDirectory = @"C:\OutputExcel";

        // Ensure the destination directory exists
        Directory.CreateDirectory(destinationDirectory);

        // Verify source directory exists
        if (!Directory.Exists(sourceDirectory))
        {
            Console.WriteLine($"Source directory not found: {sourceDirectory}");
            return;
        }

        // Get all Excel files in the source directory (xls, xlsx, xlsm)
        var excelFiles = Directory.GetFiles(sourceDirectory, "*.*", SearchOption.TopDirectoryOnly)
                                  .Where(f => f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                                              f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                                              f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase));

        foreach (var filePath in excelFiles)
        {
            try
            {
                // Ensure the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found, skipping: {filePath}");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Work with the first worksheet (adjust as needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Determine the last used row in the sheet
                int lastDataRow = sheet.Cells.MaxDataRow; // Returns -1 if the sheet is empty

                // Row where the signature will be placed (one empty row after the last data row)
                int signatureRow = (lastDataRow >= 0 ? lastDataRow + 2 : 0);

                // Insert the signature label and value
                sheet.Cells[signatureRow, 0].PutValue(SignatureLabel);
                sheet.Cells[signatureRow, 1].PutValue(SignatureValue);

                // Apply simple formatting (bold label)
                Style labelStyle = workbook.CreateStyle();
                labelStyle.Font.IsBold = true;
                sheet.Cells[signatureRow, 0].SetStyle(labelStyle);

                // Save the modified workbook to the destination folder (overwrites if exists)
                string destinationPath = Path.Combine(destinationDirectory, Path.GetFileName(filePath));
                workbook.Save(destinationPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Signature lines added to all Excel files.");
    }
}
