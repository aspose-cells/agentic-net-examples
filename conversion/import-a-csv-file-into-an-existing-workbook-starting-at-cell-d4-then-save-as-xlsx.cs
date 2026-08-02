// Title: Import CSV into Existing Workbook at D4 and Save as XLSX (C# Aspose.Cells)
// Description: Loads an existing XLSX file, imports a CSV file starting at cell D4 (row 4, column 4) with comma delimiter and numeric conversion, then saves the updated workbook as a new XLSX file. Includes file‑existence validation and error handling.
// Keywords: Aspose.Cells ImportCSV | C# import CSV into Excel | start import at D4 | save workbook as XLSX | load existing workbook Aspose | CSV to Excel conversion .NET | worksheet.Cells.ImportCSV example
// Common Searches: Aspose.Cells import CSV at specific cell | C# import CSV into existing Excel file D4 | how to merge CSV into Excel workbook using Aspose | save modified workbook as XLSX with Aspose.Cells | ImportCSV with numeric conversion C#
// Developer Intent: Load an existing workbook, import CSV data beginning at cell D4, and write the result to a new XLSX file.
// Use Cases: Populate a pre‑designed template with external CSV data at a fixed position. | Append daily CSV reports to a master spreadsheet while preserving formulas and formatting. | Combine multiple CSV sources into a single workbook without losing existing styles.
// AI Prompts: Generate C# code using Aspose.Cells to import a CSV file into cell D4 of an existing workbook and save as XLSX, including file‑existence checks. | Show how to change the start cell, delimiter, or target worksheet when importing CSV with Aspose.Cells. | Explain how to enable numeric conversion and handle import errors in Aspose.Cells ImportCSV.

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing XLSX file, imports a CSV file starting at cell D4 (row 4, column 4) with comma delimiter and numeric conversion, then saves the updated workbook as a new XLSX file. Includes file‑existence validation and error handling.
class Program
{
    static void Main()
    {
        // Paths to the existing workbook and the CSV file to be imported
        string workbookPath = "ExistingWorkbook.xlsx";
        string csvPath = "DataFile.csv";

        // Verify that the required files exist before proceeding
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Workbook file not found: {Path.GetFullPath(workbookPath)}");
            return;
        }

        if (!File.Exists(csvPath))
        {
            Console.WriteLine($"CSV file not found: {Path.GetFullPath(csvPath)}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the first worksheet (or any specific worksheet as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Import the CSV data starting at cell D4 (row index 3, column index 3)
            // Using comma as the delimiter and enabling numeric conversion.
            worksheet.Cells.ImportCSV(csvPath, ",", true, 3, 3);

            // Save the modified workbook as an XLSX file
            string resultPath = "ResultWorkbook.xlsx";
            workbook.Save(resultPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(resultPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
