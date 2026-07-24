// Title: C# – Load a comma‑delimited CSV with TxtLoadOptions and save as XLSX using Aspose.Cells
// Description: Creates a sample CSV if missing, configures TxtLoadOptions with a comma separator and numeric conversion, loads the file into an Aspose.Cells Workbook, and saves the workbook in XLSX format.
// Keywords: Aspose.Cells | TxtLoadOptions | CSV to XLSX conversion | C# spreadsheet library | comma separator | numeric data conversion | Workbook.Save | SaveFormat.Xlsx
// Common Searches: Aspose.Cells load CSV with comma separator C# | Convert CSV to XLSX using TxtLoadOptions | How to treat numeric strings as numbers when importing CSV Aspose.Cells | C# example TxtLoadOptions separator property | Save Aspose.Cells workbook as XLSX
// Developer Intent: Load a comma‑delimited CSV file with TxtLoadOptions and export it as an XLSX workbook.
// Use Cases: Automated batch conversion of CSV reports to Excel for downstream analytics. | Importing CSV data where numeric values must be recognized as numbers for financial calculations. | Generating Excel files from CSV templates in data‑processing pipelines.
// AI Prompts: Show C# code that loads a semicolon‑delimited CSV using Aspose.Cells TxtLoadOptions and saves it as XLSX. | Explain how to configure TxtLoadOptions to treat the first row of a CSV as column headers in Aspose.Cells. | Provide a step‑by‑step guide to convert multiple CSV files to XLSX in a loop with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Creates a sample CSV if missing, configures TxtLoadOptions with a comma separator and numeric conversion, loads the file into an Aspose.Cells Workbook, and saves the workbook in XLSX format.
class CsvToXlsxConverter
{
    static void Main()
    {
        // Source CSV file (comma‑delimited)
        string csvPath = "input.csv";

        // Destination XLSX file
        string xlsxPath = "output.xlsx";

        // Create a sample CSV file if it does not exist
        if (!File.Exists(csvPath))
        {
            File.WriteAllText(csvPath,
                "Name,Age,City\nJohn,30,New York\nAlice,25,London");
        }

        // Create TxtLoadOptions for CSV loading
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        loadOptions.Separator = ',';               // comma delimiter
        loadOptions.ConvertNumericData = true;     // convert numeric strings to numbers

        // Load the CSV file into a workbook using the load options
        Workbook workbook = new Workbook(csvPath, loadOptions);

        // Save the workbook as XLSX
        workbook.Save(xlsxPath, SaveFormat.Xlsx);
    }
}
