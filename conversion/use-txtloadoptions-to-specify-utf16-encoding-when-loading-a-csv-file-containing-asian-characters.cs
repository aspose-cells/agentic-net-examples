// Title: Load a CSV file with UTF‑16 (Unicode) encoding using TxtLoadOptions in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load CSV file encoded in UTF-16 | How to set TxtLoadOptions encoding to Unicode for CSV import in .NET | Reading Asian characters from a CSV using Aspose.Cells TxtLoadOptions | Specify UTF-16 encoding when loading a CSV workbook with Aspose.Cells | C# Aspose.Cells load CSV with Unicode characters without data loss
// Tags: TxtLoadOptions set UTF-16 encoding | CSV import Unicode Aspose.Cells | Load CSV containing Asian characters .NET | Aspose.Cells workbook load options for text files | C# specify encoding for CSV import

using System;
using System.Text;
using Aspose.Cells;

// The example creates a TxtLoadOptions object for CSV format, sets its Encoding property to Encoding.Unicode (UTF‑16), and loads the CSV file into an Aspose.Cells Workbook. It demonstrates handling Asian characters correctly and optionally saving the workbook as XLSX.
class LoadCsvWithUtf16
{
    static void Main()
    {
        // Specify the path to the CSV file
        string csvPath = "input.csv";

        // Create TxtLoadOptions for CSV format
        TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.CSV);

        // Set the encoding to UTF‑16 (Unicode)
        loadOptions.Encoding = Encoding.Unicode; // UTF-16

        // Load the CSV file using the specified options
        Workbook workbook = new Workbook(csvPath, loadOptions);

        // (Optional) Save the workbook to another format, e.g., XLSX
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
