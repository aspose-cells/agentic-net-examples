// Title: Convert an Excel workbook to a gzipped CSV file using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# export first worksheet to CSV and gzip the file | How to save an .xlsx as .csv.gz with Aspose.Cells and .NET GZipStream | C# code to convert Excel to compressed CSV using TxtSaveOptions | Create a gzipped CSV from Excel workbook with Aspose.Cells in .NET
// Tags: Aspose.Cells CSV export TxtSaveOptions | GZipStream compress CSV in C# | Excel to gzipped CSV conversion using Aspose.Cells | temporary CSV file cleanup after compression | SaveFormat.CSV usage with Aspose.Cells

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

// The example loads an Excel workbook, saves the first worksheet as a CSV file using TxtSaveOptions, compresses the CSV into a GZIP archive with GZipStream, and then deletes the intermediate CSV file.
class Program
{
    static void Main(string[] args)
    {
        // Path to the source Excel workbook
        string inputPath = "input.xlsx";

        // Path for the intermediate CSV file
        string csvPath = "output.csv";

        // Path for the final GZIP-compressed CSV
        string gzipPath = "output.csv.gz";

        // Load the workbook from the Excel file
        Workbook workbook = new Workbook(inputPath);

        // Save the first worksheet as CSV
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV);
        // (Optional) configure CSV options here, e.g., csvOptions.Encoding = Encoding.UTF8;
        workbook.Save(csvPath, csvOptions);

        // Compress the CSV file into a GZIP archive
        using (FileStream csvFile = new FileStream(csvPath, FileMode.Open, FileAccess.Read))
        using (FileStream gzipFile = new FileStream(gzipPath, FileMode.Create, FileAccess.Write))
        using (GZipStream gzipStream = new GZipStream(gzipFile, CompressionMode.Compress))
        {
            csvFile.CopyTo(gzipStream);
        }

        // Clean up the temporary CSV file
        File.Delete(csvPath);
    }
}
