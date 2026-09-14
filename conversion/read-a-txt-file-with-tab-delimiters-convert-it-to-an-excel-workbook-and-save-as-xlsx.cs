// Title: Convert a tab‑delimited TXT (TSV) file to an XLSX workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells ConversionUtility to read a TSV file with LoadOptions(LoadFormat.Tsv) and write it as an XLSX workbook via OoxmlSaveOptions in C#. | Add try‑catch error handling around the Aspose.Cells conversion of a tab‑delimited text file to Excel. | Specify source and destination paths and perform a format conversion from TSV to XLSX with Aspose.Cells in a console application.
// Common Searches: aspnet convert tab delimited txt file to xlsx using aspose.cells c# example | c# load tsv file with Aspose.Cells LoadOptions and save as xlsx | how to use Aspose.Cells ConversionUtility to change file format from TSV to Excel | sample code for converting a TSV text file to an OOXML workbook in .NET
// Tags: TSV to XLSX conversion Aspose.Cells | Aspose.Cells LoadOptions for tab delimited files | Aspose.Cells OoxmlSaveOptions XLSX output | ConversionUtility format conversion C# | exception handling Aspose.Cells file conversion

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    // // Demonstrates reading a tab‑delimited (TSV) text file with LoadOptions(LoadFormat.Tsv), converting it to an Excel workbook, and saving the result as an XLSX file using OoxmlSaveOptions via Aspose.Cells ConversionUtility, including basic exception handling.
    class Program
    {
        static void Main()
        {
            // Path to the source tab‑delimited text file
            string sourcePath = "input.txt";

            // Desired path for the resulting Excel workbook
            string destinationPath = "output.xlsx";

            try
            {
                // Load options specifying that the source file is a tab‑delimited (TSV) file
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);

                // Save options for XLSX (OOXML) format
                SaveOptions saveOptions = new OoxmlSaveOptions();

                // Perform the conversion using the provided ConversionUtility method
                ConversionUtility.Convert(sourcePath, loadOptions, destinationPath, saveOptions);

                Console.WriteLine($"Conversion successful: \"{sourcePath}\" → \"{destinationPath}\"");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}
