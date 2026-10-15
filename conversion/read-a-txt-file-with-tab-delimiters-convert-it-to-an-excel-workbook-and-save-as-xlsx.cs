// Title: Convert a tab‑delimited TXT file to an XLSX workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load tab delimited text file and save as xlsx | How to configure TxtLoadOptions separator for TSV in Aspose.Cells | Convert numeric and date strings from a .txt file to Excel with Aspose.Cells | Example of reading a tab‑separated .txt and exporting to .xlsx using Aspose.Cells .NET | Aspose.Cells code sample for TSV to Excel conversion in C#
// Tags: TxtLoadOptions separator configuration | Aspose.Cells import TSV to Workbook | Workbook.Save Xlsx format Aspose.Cells | automatic numeric conversion TxtLoadOptions | date time parsing TxtLoadOptions Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program uses Aspose.Cells for .NET to load a tab‑delimited TXT (TSV) file with TxtLoadOptions (setting the tab character as the separator and enabling numeric/date conversion), creates a Workbook from the data, and saves the result as an XLSX file.
class TxtToExcelConverter
{
    static void Main()
    {
        // Path to the source tab‑delimited text file
        string txtFilePath = "input.txt";

        // Path for the resulting Excel workbook
        string excelFilePath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(txtFilePath))
                throw new FileNotFoundException($"Input file not found: {txtFilePath}");

            // Configure load options for a tab‑delimited TXT file
            TxtLoadOptions loadOptions = new TxtLoadOptions(); // default constructor
            loadOptions.Separator = '\t';               // Use tab as the column separator
            loadOptions.ConvertNumericData = true;      // Convert numeric strings to numbers
            loadOptions.ConvertDateTimeData = true;     // Convert date strings to DateTime

            // Load the TXT file into a workbook using the specified options
            Workbook workbook = new Workbook(txtFilePath, loadOptions);

            // Save the workbook as an XLSX file
            workbook.Save(excelFilePath, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion successful. Excel file saved to: {excelFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
        }
    }
}
