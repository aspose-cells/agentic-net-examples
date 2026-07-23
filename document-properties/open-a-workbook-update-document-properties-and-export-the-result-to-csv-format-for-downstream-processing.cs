// Title: C# – Update Excel built‑in & custom document properties and export to CSV with Aspose.Cells
// Description: Loads an Excel workbook, changes the built‑in Author property, adds a custom ProcessedDate property, and saves the modified file directly as CSV for downstream processing using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# document properties | set Excel Author property Aspose.Cells | add custom Excel property C# | export Excel to CSV Aspose.Cells | modify Excel metadata .NET | SaveFormat.Csv Aspose.Cells
// Common Searches: How to change Author property in Excel with Aspose.Cells C# | Add custom property to Excel workbook and save as CSV using Aspose | Aspose.Cells update document properties before CSV conversion | C# code to set Excel metadata and export to CSV | Aspose.Cells document properties example
// Developer Intent: Update built‑in and custom document properties of an Excel file and generate a CSV version in one workflow.
// Use Cases: Embed author information before sending the file to a review system that reads metadata. | Record processing timestamps as custom properties for audit trails, then convert to CSV for data pipelines. | Automate batch processing of Excel reports to standardize metadata and produce CSV outputs for analytics.
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, updates the Author built‑in property, adds a custom ProcessedDate property, and saves the workbook as CSV. | Create a reusable method that accepts an input Excel path, author name, and custom property value, updates the document properties using Aspose.Cells, and returns the CSV as a byte array. | Explain error‑handling strategies when modifying document properties and exporting to CSV with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Loads an Excel workbook, changes the built‑in Author property, adds a custom ProcessedDate property, and saves the modified file directly as CSV for downstream processing using Aspose.Cells for .NET.
class WorkbookCsvExport
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Load the workbook from the file
        using (Workbook workbook = new Workbook(sourcePath))
        {
            // Update a built‑in document property (Author)
            DocumentProperty authorProp = workbook.BuiltInDocumentProperties["Author"];
            authorProp.Value = "John Doe";

            // Add a custom document property (ProcessedDate)
            workbook.CustomDocumentProperties.Add("ProcessedDate", DateTime.Now);

            // Export the workbook to CSV format
            string csvPath = "output.csv";
            workbook.Save(csvPath, SaveFormat.Csv);
        }

        Console.WriteLine("Workbook properties updated and saved as CSV.");
    }
}
