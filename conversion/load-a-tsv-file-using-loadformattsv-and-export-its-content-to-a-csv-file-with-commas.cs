// Title: How to convert a TSV file to a comma‑separated CSV with Aspose.Cells in C# (.NET)
// AI Prompts: Load a tab‑delimited file using Aspose.Cells TxtLoadOptions with a '\t' separator and then save it as a CSV using TxtSaveOptions with a ',' separator. | Generate C# code that reads an input.tsv, sets TxtLoadOptions.Separator to tab, and writes output.csv with TxtSaveOptions.Separator set to comma.
// Common Searches: aspocells c# convert tab separated values file to csv | load tsv with TxtLoadOptions and export as comma delimited csv example | c# Aspose.Cells save workbook as csv with custom delimiter | how to set separator for TxtLoadOptions in Aspose.Cells | convert large tsv to csv using Aspose.Cells .NET
// Tags: tab‑separator loading options Aspose.Cells | comma‑separator saving options Aspose.Cells | TSV to CSV conversion using Aspose.Cells | C# read TSV with Aspose.Cells | export workbook to CSV Aspose.Cells

using System;
using Aspose.Cells;

// // Loads a tab‑separated values (TSV) file via TxtLoadOptions (tab separator) and saves it as a comma‑separated CSV using TxtSaveOptions (comma separator).
class TsvToCsvConverter
{
    static void Main()
    {
        // Input TSV file path
        string inputPath = "input.tsv";

        // Output CSV file path
        string outputPath = "output.csv";

        // Load options for TSV format
        TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Tsv);
        // TSV uses tab as separator; set explicitly for clarity
        loadOptions.Separator = '\t';

        // Load the TSV file into a workbook
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Save options for CSV with comma separator
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        saveOptions.Separator = ',';
        // Optional: set encoding if needed
        // saveOptions.Encoding = System.Text.Encoding.UTF8;

        // Save the workbook as a CSV file
        workbook.Save(outputPath, saveOptions);
    }
}
