// Title: C# – Convert Tab‑Delimited TXT to XLSX with Aspose.Cells
// Description: Shows how to read a TSV (tab‑separated) text file using Aspose.Cells LoadOptions (LoadFormat.Tsv) and convert it to an Excel workbook (XLSX) via ConversionUtility and OoxmlSaveOptions.
// Keywords: Aspose.Cells | C# TSV to XLSX conversion | LoadFormat.Tsv | ConversionUtility example | OoxmlSaveOptions | tab delimited to Excel | TSV to Excel .NET | GitHub Aspose.Cells sample
// Common Searches: Aspose.Cells convert txt to xlsx c# | C# load tab delimited file into Excel | ConversionUtility TSV example | LoadOptions Tsv Aspose.Cells usage | Save workbook as xlsx using OoxmlSaveOptions
// Developer Intent: Programmatically transform a tab‑separated .txt file into an .xlsx workbook using Aspose.Cells for .NET.
// Use Cases: Automate nightly import of TSV reports into Excel for analytics. | Batch‑process multiple data‑dump .txt files into user‑friendly XLSX workbooks. | Embed TSV‑to‑Excel conversion in a data‑pipeline that feeds downstream BI tools.
// AI Prompts: Write C# code that reads a tab‑delimited .txt file and converts it to .xlsx with Aspose.Cells, including comprehensive error handling and logging. | Explain the role of LoadOptions (LoadFormat.Tsv) and OoxmlSaveOptions when converting a TSV file to an Excel workbook in Aspose.Cells. | Provide a C# script that scans a folder for .txt files, converts each to .xlsx using ConversionUtility, and records success/failure in a CSV log.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace TxtToExcelConversion
{
    // Shows how to read a TSV (tab‑separated) text file using Aspose.Cells LoadOptions (LoadFormat.Tsv) and convert it to an Excel workbook (XLSX) via ConversionUtility and OoxmlSaveOptions.
    class Program
    {
        static void Main()
        {
            // Path to the source tab‑delimited text file
            string sourcePath = "data.txt";

            // Desired output Excel file path
            string outputPath = "result.xlsx";

            // Load options specifying that the source file is a TSV (tab‑separated) file
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);

            // Save options for creating an XLSX workbook
            SaveOptions saveOptions = new OoxmlSaveOptions();

            // Perform the conversion using Aspose.Cells.Utility.ConversionUtility
            ConversionUtility.Convert(sourcePath, loadOptions, outputPath, saveOptions);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{outputPath}'");
        }
    }
}
