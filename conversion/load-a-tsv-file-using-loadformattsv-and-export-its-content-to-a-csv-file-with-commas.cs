// Title: Load a TSV file and export it to CSV using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# load TSV file with LoadFormat.Tsv and save as CSV | Convert tab‑separated values to comma‑separated values using Aspose.Cells in .NET | C# example for TxtSaveOptions SaveFormat.CSV after loading a TSV workbook | How to read a .tsv file with Aspose.Cells and write a .csv file in C# | Using LoadOptions to import TSV and TxtSaveOptions to export CSV in .NET
// Tags: LoadOptions with LoadFormat.Tsv | TxtSaveOptions for CSV output | TSV to CSV conversion using Aspose.Cells | Workbook.Save CSV format C# example | Aspose.Cells import TSV file

using Aspose.Cells;
using System;

// The sample loads a tab‑separated values (TSV) file using LoadOptions configured with LoadFormat.Tsv, then saves the workbook as a comma‑separated values (CSV) file using TxtSaveOptions set to SaveFormat.CSV.
class Program
{
    static void Main()
    {
        // Load the TSV file using the TSV load format
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);
        Workbook workbook = new Workbook("input.tsv", loadOptions);

        // Prepare CSV save options (comma is the default delimiter for CSV)
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.CSV);

        // Export the workbook content to a CSV file
        workbook.Save("output.csv", saveOptions);
    }
}
