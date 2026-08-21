// Title: C# – Import CSV into an Existing Workbook at D4 and Save as XLSX with Aspose.Cells
// Description: Loads an existing Excel file, imports data from a CSV file starting at cell D4 (row 4, column 4) using a comma delimiter and automatic numeric conversion, then saves the workbook as an XLSX file.
// Keywords: Aspose.Cells | C# import CSV | ImportCSV | start cell D4 | save as XLSX | load existing workbook | CSV to Excel conversion | Aspose.Cells example | cells.ImportCSV | Excel template merge
// Common Searches: Aspose.Cells import CSV at specific cell | C# import CSV into existing Excel workbook | How to start CSV import at D4 using Aspose.Cells | Save workbook as XLSX after CSV import C# | Aspose.Cells Cells.ImportCSV parameters
// Developer Intent: Load a workbook, insert CSV data beginning at D4, and export the result as an XLSX file.
// Use Cases: Combine a CSV export with a pre‑formatted Excel template without overwriting existing formulas or styles. | Append periodic CSV reports into a master spreadsheet by inserting them at a designated location. | Automate the creation of a new XLSX report by merging raw CSV data with existing worksheets that contain branding and calculations.
// AI Prompts: Write C# code that uses Aspose.Cells to load an existing workbook, import a CSV file starting at cell D4 with numeric conversion, and save the result as XLSX. | Explain how to change the delimiter, start row, and start column when calling Cells.ImportCSV in Aspose.Cells. | Provide best‑practice error handling for loading a workbook, importing CSV data, and saving the file with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsCsvImportDemo
{
    // Loads an existing Excel file, imports data from a CSV file starting at cell D4 (row 4, column 4) using a comma delimiter and automatic numeric conversion, then saves the workbook as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Path to the existing workbook that will receive the CSV data
            string existingWorkbookPath = "existing.xlsx";

            // Load the existing workbook (lifecycle rule: Workbook(string))
            Workbook workbook = new Workbook(existingWorkbookPath);

            // Access the first worksheet (you can change the index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Path to the CSV file to be imported
            string csvFilePath = "data.csv";

            // Import CSV starting at cell D4 (row index 3, column index 3)
            // Using comma as the delimiter and converting numeric strings to numbers
            cells.ImportCSV(csvFilePath, ",", true, 3, 3);

            // Save the modified workbook as XLSX (lifecycle rule: Workbook.Save)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"CSV data imported into '{existingWorkbookPath}' at D4 and saved as '{outputPath}'.");
        }
    }
}
