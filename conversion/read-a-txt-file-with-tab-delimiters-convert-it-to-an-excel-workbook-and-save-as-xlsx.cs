// Title: C# – Convert Tab‑Delimited TXT to XLSX with Aspose.Cells
// Description: Shows how to load a tab‑separated text (TSV) file using Aspose.Cells TxtLoadOptions, create a Workbook, and save it as an XLSX workbook in .NET.
// Keywords: Aspose.Cells | C# | TSV to XLSX | tab delimited txt to excel | TxtLoadOptions | LoadFormat.Tsv | Excel conversion .NET | save as xlsx | text to excel conversion
// Common Searches: Aspose.Cells load TSV file C# | convert txt tab delimited to xlsx .NET | TxtLoadOptions separator tab | C# read tab separated values into Excel | save workbook as xlsx using Aspose.Cells
// Developer Intent: Read a tab‑delimited .txt file into an Aspose.Cells Workbook and export it as an .xlsx file.
// Use Cases: Automate nightly conversion of TSV reports for business analysts. | Provide a utility that transforms exported log files (tab‑separated) into .xlsx for easy sharing. | Integrate TSV‑to‑XLSX conversion into a data‑import pipeline using Aspose.Cells.
// AI Prompts: Generate C# code that reads a tab‑delimited .txt file with Aspose.Cells TxtLoadOptions and saves it as .xlsx, including basic error handling. | Provide an example of converting a TSV file to Excel with Aspose.Cells, showing how to set the separator and choose the output format. | Explain how to configure TxtLoadOptions for various delimiters and save the resulting workbook in multiple Excel formats.

using System;
using Aspose.Cells;

namespace TxtToExcelConversion
{
    // Shows how to load a tab‑separated text (TSV) file using Aspose.Cells TxtLoadOptions, create a Workbook, and save it as an XLSX workbook in .NET.
    class Program
    {
        static void Main()
        {
            // Path to the source tab‑delimited text file
            string txtPath = "input.txt";

            // Path for the resulting Excel workbook
            string xlsxPath = "output.xlsx";

            // Configure load options for a tab‑delimited (TSV) file
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Tsv);
            loadOptions.Separator = '\t'; // Tab character as delimiter

            // Load the text file into a workbook using the specified options
            Workbook workbook = new Workbook(txtPath, loadOptions);

            // Save the workbook as an XLSX file
            workbook.Save(xlsxPath, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion completed: \"{txtPath}\" → \"{xlsxPath}\"");
        }
    }
}
