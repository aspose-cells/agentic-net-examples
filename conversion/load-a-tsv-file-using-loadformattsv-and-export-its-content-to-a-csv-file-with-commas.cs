// Title: C# – Convert TSV to CSV Using Aspose.Cells (LoadFormat.Tsv → SaveFormat.Csv)
// Description: Shows how to load a tab‑separated values (TSV) file with TxtLoadOptions (LoadFormat.Tsv) into an Aspose.Cells Workbook and export it as a comma‑separated CSV using TxtSaveOptions (SaveFormat.Csv).
// Keywords: Aspose.Cells TSV to CSV | LoadFormat.Tsv C# | TxtLoadOptions example | TxtSaveOptions CSV | Aspose.Cells conversion .NET | tab separated values Aspose.Cells | comma separated CSV Aspose.Cells | C# spreadsheet format conversion | Aspose.Cells CSV export | global .NET data migration
// Common Searches: Aspose.Cells load tsv file | convert tsv to csv C# Aspose.Cells | TxtLoadOptions LoadFormat.Tsv usage | save workbook as csv Aspose.Cells | change separator Aspose.Cells TxtSaveOptions | C# example TSV to CSV Aspose.Cells | Aspose.Cells data format conversion tutorial
// Developer Intent: Load a TSV file and write its contents to a comma‑separated CSV file using Aspose.Cells for .NET.
// Use Cases: Migrate legacy tab‑delimited reports to CSV for compatibility with modern spreadsheet tools. | Prepare data exports for third‑party systems that accept only CSV input. | Automate batch conversion of multiple TSV datasets in a .NET processing pipeline.
// AI Prompts: Generate C# code that reads a large TSV file with Aspose.Cells and streams it to a CSV with commas, including memory‑efficient handling. | Explain how to configure TxtLoadOptions and TxtSaveOptions for custom delimiters when converting between TSV and CSV in Aspose.Cells. | Provide error‑handling best practices for file I/O and format mismatches during TSV‑to‑CSV conversion with Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to load a tab‑separated values (TSV) file with TxtLoadOptions (LoadFormat.Tsv) into an Aspose.Cells Workbook and export it as a comma‑separated CSV using TxtSaveOptions (SaveFormat.Csv).
class TsvToCsvConverter
{
    static void Main()
    {
        // Input TSV file path
        string tsvPath = "input.tsv";

        // Output CSV file path
        string csvPath = "output.csv";

        // Load options for TSV format (tab‑separated)
        TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Tsv);
        loadOptions.Separator = '\t'; // Explicitly set tab as separator

        // Load the TSV file into a workbook
        Workbook workbook = new Workbook(tsvPath, loadOptions);

        // Save options for CSV format with comma separator
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv);
        saveOptions.Separator = ','; // Use comma for CSV output

        // Export the workbook content to CSV
        workbook.Save(csvPath, saveOptions);
    }
}
