// Title: C# – Update Excel Built‑in & Custom Document Properties and Save as CSV using Aspose.Cells
// Description: Load an existing workbook with Aspose.Cells for .NET, modify the Author and Title built‑in properties, add a custom boolean flag, and export the result directly to CSV for downstream processing.
// Keywords: Aspose.Cells C# update document properties | set Excel author title .NET | add custom Excel property Aspose | export workbook to CSV Aspose.Cells | save Excel as CSV C# | modify built‑in document properties | custom document property boolean | CSV conversion Aspose.Cells | programmatic Excel metadata | Aspose.Cells SaveFormat.Csv
// Common Searches: Aspose.Cells change Excel author property C# | How to add custom document property in Aspose.Cells | Export Excel to CSV after updating metadata with Aspose | C# code to set built‑in properties and save as CSV | Aspose.Cells document properties example | Convert Excel to CSV using Aspose.Cells after editing properties
// Developer Intent: Programmatically modify workbook metadata and convert the file to CSV.
// Use Cases: Generate CSV reports that include standardized author and title metadata for compliance audits. | Mark processed workbooks with a custom flag before bulk CSV conversion in ETL pipelines. | Enforce document governance by setting built‑in properties prior to exporting data to CSV. | Add a preprocessing step that enriches Excel files with metadata before downstream analytics. | Track processing status with a boolean property when converting Excel files to CSV.
// AI Prompts: Provide C# Aspose.Cells code that sets the Author and Title built‑in properties, adds a boolean custom property named Processed, and saves the workbook as a CSV file. | Show how to safely add or update a custom document property in an Excel workbook using Aspose.Cells, then export it to CSV with a specific encoding. | Explain the steps to modify both built‑in and custom document properties in Aspose.Cells before converting the workbook to CSV, including error handling for existing properties.

using System;
using Aspose.Cells;

// Load an existing workbook with Aspose.Cells for .NET, modify the Author and Title built‑in properties, add a custom boolean flag, and export the result directly to CSV for downstream processing.
class Program
{
    static void Main()
    {
        // Path to the source Excel workbook
        string sourcePath = "input.xlsx";

        // Path where the CSV file will be saved
        string csvPath = "output.csv";

        // Load the workbook from the file system
        Workbook workbook = new Workbook(sourcePath);

        // Update built‑in document properties
        workbook.BuiltInDocumentProperties["Author"].Value = "John Doe";
        workbook.BuiltInDocumentProperties["Title"].Value = "Sales Report";

        // Add a custom document property (if it already exists, this will throw;
        // for simplicity we assume it does not exist)
        workbook.CustomDocumentProperties.Add("Processed", true);

        // Export the workbook to CSV format
        workbook.Save(csvPath, SaveFormat.Csv);
    }
}
