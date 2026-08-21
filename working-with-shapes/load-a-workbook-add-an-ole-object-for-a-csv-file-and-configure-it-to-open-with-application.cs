// Title: Embed CSV as OLE Object in Excel Workbook using Aspose.Cells for .NET
// Description: Loads an existing .xlsx file, reads a CSV into a byte array, creates a tiny transparent PNG placeholder, adds an OLE object to the first worksheet, embeds the CSV (non‑linked), sets ProgID to Excel.CSV, enables AutoLoad, and saves the result.
// Keywords: Aspose.Cells | C# OLE object | embed CSV in Excel | ProgID Excel.CSV | AutoLoad OLE | placeholder image | .NET Excel automation | OleObject Add | CSV embedded workbook
// Common Searches: How to embed a CSV file as an OLE object with Aspose.Cells .NET | Aspose.Cells set ProgID to Excel.CSV for OLE objects | C# add OLE object to worksheet and auto‑load CSV | Embedding non‑linked CSV in Excel using Aspose.Cells | Create Excel file with embedded CSV using Aspose.Cells
// Developer Intent: Add a CSV file as an embedded OLE object that opens automatically in Excel.
// Use Cases: Generate a report workbook that contains a CSV data source for quick viewing or editing. | Distribute a single Excel file that carries its CSV attachment, eliminating separate downloads. | Automate archival of CSV datasets inside Excel files for compliance or data‑exchange workflows.
// AI Prompts: Write C# code with Aspose.Cells to embed a CSV as an OLE object, using a transparent PNG placeholder and setting ProgID to Excel.CSV. | Show how to add an OLE object to a worksheet, embed binary CSV data, enable AutoLoad, and save the workbook in Aspose.Cells. | Explain the steps to configure an OLE object in Aspose.Cells so the embedded CSV opens automatically when the workbook is opened.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Required for OleObject

namespace AsposeCellsOleObjectDemo
{
    // Loads an existing .xlsx file, reads a CSV into a byte array, creates a tiny transparent PNG placeholder, adds an OLE object to the first worksheet, embeds the CSV (non‑linked), sets ProgID to Excel.CSV, enables AutoLoad, and saves the result.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Input workbook path
                string workbookPath = "input.xlsx";
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // First worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // CSV file to embed
                string csvFilePath = "data.csv";
                if (!File.Exists(csvFilePath))
                {
                    Console.WriteLine($"CSV file not found: {csvFilePath}");
                    return;
                }

                // Read CSV bytes
                byte[] csvData = File.ReadAllBytes(csvFilePath);

                // Minimal 1x1 PNG placeholder image (transparent)
                byte[] placeholderImage = Convert.FromBase64String(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK5cAAAAASUVORK5CYII=");

                // Add OLE object with placeholder image
                int oleIndex = sheet.OleObjects.Add(5, 5, 200, 200, placeholderImage);
                OleObject ole = sheet.OleObjects[oleIndex];

                // Embed CSV data (not linked)
                ole.SetEmbeddedObject(false, csvData, Path.GetFileName(csvFilePath), false, "CSV Data");

                // Set ProgID so Excel opens it as CSV
                ole.ProgID = "Excel.CSV";

                // Auto load when workbook opens
                ole.AutoLoad = true;

                // Save the modified workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved with embedded CSV OLE object at '{outputPath}'.");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
