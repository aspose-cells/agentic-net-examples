using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Paths for the source TSV file and the destination CSV file
        string tsvPath = "input.tsv";
        string csvPath = "output.csv";

        // Create load options for a TSV file (tab‑separated)
        TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Tsv);
        loadOptions.Separator = '\t'; // Explicitly set tab as the separator

        // Load the TSV file into a workbook using the specified options
        Workbook workbook = new Workbook(tsvPath, loadOptions);

        // Create save options for CSV format (comma‑separated)
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv);
        saveOptions.Separator = ','; // Ensure commas are used as delimiters

        // Save the workbook content as a CSV file
        workbook.Save(csvPath, saveOptions);
    }
}