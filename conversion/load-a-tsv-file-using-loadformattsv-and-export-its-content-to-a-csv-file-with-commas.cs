// Title: Convert TSV to CSV with Aspose.Cells for .NET (C#) using TxtLoadOptions & TxtSaveOptions
// Description: Load a tab‑separated file with Aspose.Cells TxtLoadOptions (LoadFormat.Tsv) and save it as a comma‑separated CSV using TxtSaveOptions. Compact C# sample.
// Keywords: Aspose.Cells TSV to CSV | C# TxtLoadOptions LoadFormat.Tsv | Aspose.Cells TxtSaveOptions CSV | convert .tsv to .csv .NET | Aspose.Cells delimiter conversion | load TSV Aspose.Cells | save CSV Aspose.Cells
// Common Searches: Aspose.Cells load TSV file C# | convert tsv to csv using Aspose.Cells | TxtLoadOptions LoadFormat.Tsv example | Aspose.Cells export CSV with custom separator | C# code to change delimiter with Aspose.Cells
// Developer Intent: Read a TSV file and write its data to a CSV file with commas using Aspose.Cells for .NET.
// Use Cases: Batch conversion of tab‑delimited reports to CSV for analytics pipelines. | Preparing data exports from legacy systems (TSV) for tools that require CSV input. | Automating file format migration in ETL workflows where source files are TSV and target systems expect CSV.
// AI Prompts: Generate C# code that loads a .tsv file with Aspose.Cells TxtLoadOptions (LoadFormat.Tsv) and saves it as a .csv file using TxtSaveOptions with a comma separator. | Explain how to switch delimiters when loading and saving text files with Aspose.Cells, including handling of custom separators. | Provide robust error handling for missing or malformed TSV files during conversion to CSV with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsTsvToCsv
{
    // Load a tab‑separated file with Aspose.Cells TxtLoadOptions (LoadFormat.Tsv) and save it as a comma‑separated CSV using TxtSaveOptions. Compact C# sample.
    class Program
    {
        static void Main()
        {
            // Paths for the source TSV file and the destination CSV file
            string tsvFilePath = "input.tsv";
            string csvFilePath = "output.csv";

            // Load the TSV file using TxtLoadOptions with LoadFormat.Tsv
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Tsv);
            // Explicitly set the tab character as the separator (optional, default for TSV)
            loadOptions.Separator = '\t';

            // Create a workbook from the TSV file with the specified load options
            Workbook workbook = new Workbook(tsvFilePath, loadOptions);

            // Prepare save options for CSV output with comma as the separator
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv);
            saveOptions.Separator = ',';

            // Save the workbook content as a CSV file using the configured save options
            workbook.Save(csvFilePath, saveOptions);

            Console.WriteLine($"TSV file '{tsvFilePath}' has been converted to CSV file '{csvFilePath}'.");
        }
    }
}
