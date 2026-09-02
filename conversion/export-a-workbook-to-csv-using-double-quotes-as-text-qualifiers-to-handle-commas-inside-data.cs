// Title: Export an Aspose.Cells workbook to a CSV file with all fields quoted using double‑quote text qualifiers in C#
// AI Prompts: Generate C# code that creates a workbook, adds data containing commas, and saves it as a CSV where every cell value is wrapped in double quotes using Aspose.Cells. | Show how to configure TxtSaveOptions in Aspose.Cells to use a comma separator and always apply the double‑quote text qualifier when exporting to CSV. | Provide a complete example that demonstrates producing a quoted CSV suitable for systems that require text qualifiers, using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# export workbook to CSV with always quoted fields | How to set TxtSaveOptions to force double quotes around values in CSV output | Saving Excel data containing commas to CSV using Aspose.Cells and text qualifiers | C# Aspose.Cells CSV export with QuoteType.Always and comma delimiter
// Tags: TxtSaveOptions CSV text qualifier | Aspose.Cells always quote CSV values | C# generate quoted CSV with Aspose.Cells | CSV export handling commas Aspose.Cells | QuoteType.Always configuration Aspose.Cells

using System;
using Aspose.Cells;

// Demonstrates creating a workbook, inserting data that includes commas, and saving it as a CSV file where every field is enclosed in double quotes by configuring TxtSaveOptions with a comma separator and QuoteType.Always.
class ExportWorkbookToCsv
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add sample data that contains commas
        cells["A1"].PutValue("Name");
        cells["B1"].PutValue("Description");
        cells["A2"].PutValue("John Doe");
        cells["B2"].PutValue("Engineer, Software");
        cells["A3"].PutValue("Jane Smith");
        cells["B3"].PutValue("Manager, Sales");

        // Configure CSV save options to use double quotes as text qualifiers
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv);
        saveOptions.Separator = ',';                     // Use comma as delimiter
        saveOptions.QuoteType = TxtValueQuoteType.Always; // Always enclose fields in double quotes

        // Save the workbook as a CSV file
        string outputPath = "output.csv";
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine($"Workbook exported to CSV with double quotes at '{outputPath}'.");
    }
}
