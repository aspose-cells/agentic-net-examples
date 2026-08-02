// Title: Batch convert CSV files to ODS with Aspose.Cells for .NET (default ODF version) and verify each workbook
// Description: C# example that scans a folder, loads every *.csv using Aspose.Cells LoadOptions, converts each file to an ODS workbook with default OdsSaveOptions via ConversionUtility, saves it with a .ods extension, then reloads the ODS to confirm worksheet count and logs success or errors.
// Keywords: Aspose.Cells | C# CSV to ODS | batch conversion .NET | default ODF version | OdsSaveOptions | ConversionUtility | verify ODS output | load CSV directory | LibreOffice compatibility | OpenOffice ODS | GitHub example | global developers | US developers | Europe developers
// Common Searches: C# batch convert CSV to ODS Aspose.Cells | How to verify ODS after conversion .NET | Default ODF version when saving ODS with Aspose | Convert all CSV files in a folder to ODS using Aspose | Aspose.Cells example for bulk CSV to ODS conversion
// Developer Intent: Automatically transform every CSV file in a specified directory into an ODS spreadsheet using Aspose.Cells’s default ODF settings and ensure each output file is structurally valid.
// Use Cases: Nightly automation that turns exported CSV reports into ODS files for LibreOffice or OpenOffice consumption. | Bulk data‑import pipelines where source CSVs must be packaged as ODS workbooks before downstream analysis. | Quality‑control scripts that confirm each generated ODS contains the expected number of worksheets and logs any failures.
// AI Prompts: Generate C# code that iterates through a folder of CSV files, uses Aspose.Cells ConversionUtility to convert each to ODS with default OdsSaveOptions, then reloads each ODS to verify worksheet count. | Provide a robust error‑handling and detailed logging pattern for batch CSV‑to‑ODS conversion with Aspose.Cells. | Show how to customize OdsSaveOptions to target a specific ODF version while still performing post‑conversion verification of the ODS workbooks.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// C# example that scans a folder, loads every *.csv using Aspose.Cells LoadOptions, converts each file to an ODS workbook with default OdsSaveOptions via ConversionUtility, saves it with a .ods extension, then reloads the ODS to confirm worksheet count and logs success or errors.
class CsvToOdsBatch
{
    static void Main()
    {
        // Directory containing CSV files
        string csvDirectory = "CsvFiles";

        // Verify the directory exists
        if (!Directory.Exists(csvDirectory))
        {
            Console.WriteLine($"Directory not found: {csvDirectory}");
            return;
        }

        // Get all CSV files in the directory
        string[] csvFiles = Directory.GetFiles(csvDirectory, "*.csv");

        foreach (string csvFilePath in csvFiles)
        {
            try
            {
                // Destination ODS file path (same name, .ods extension)
                string odsFilePath = Path.ChangeExtension(csvFilePath, ".ods");

                // Load options for CSV files
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);

                // Save options for ODS (default ODF version)
                OdsSaveOptions saveOptions = new OdsSaveOptions();

                // Convert CSV to ODS using the utility method
                ConversionUtility.Convert(csvFilePath, loadOptions, odsFilePath, saveOptions);
                Console.WriteLine($"Converted: {Path.GetFileName(csvFilePath)} → {Path.GetFileName(odsFilePath)}");

                // Verify compliance by loading the generated ODS file
                OdsLoadOptions odsLoadOptions = new OdsLoadOptions();
                Workbook workbook = new Workbook(odsFilePath, odsLoadOptions);
                Console.WriteLine($"Verification: {workbook.Worksheets.Count} worksheet(s) in {Path.GetFileName(odsFilePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing {Path.GetFileName(csvFilePath)}: {ex.Message}");
            }
        }
    }
}
