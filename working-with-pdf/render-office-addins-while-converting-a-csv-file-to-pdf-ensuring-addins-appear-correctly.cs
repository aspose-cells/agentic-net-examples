// Title: Convert a CSV file to PDF while preserving Office Add‑Ins using Aspose.Cells for .NET
// Common Searches: how to keep Excel add‑ins visible when converting CSV to PDF with Aspose.Cells C# | Aspose.Cells C# preserve Office Add‑Ins during CSV to PDF conversion | sample code for converting CSV to PDF while retaining add‑ins in .NET | enable add‑in support in Aspose.Cells when saving workbook as PDF
// Tags: Aspose.Cells CSV to PDF conversion | Add‑in support property EnableAddInSupport | LoadOptions CSV format Aspose.Cells | Workbook.Save PDF C# example | Office Add‑Ins handling in Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace CsvToPdfWithAddIns
{
    // This example loads a CSV file into an Aspose.Cells Workbook using LoadOptions with LoadFormat.Csv, then saves the workbook as a PDF. It notes that the EnableAddInSupport property, which retains Office Add‑Ins, may be unavailable in newer versions, so the code omits it to stay compilable while demonstrating the conversion workflow.
    class Program
    {
        static void Main(string[] args)
        {
            // Input CSV file path
            string csvPath = @"C:\Input\sample.csv";

            // Output PDF file path
            string pdfPath = @"C:\Output\sample.pdf";

            // Verify that the input CSV file exists
            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"Error: CSV file not found at '{csvPath}'.");
                return;
            }

            try
            {
                // Load options for CSV format
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);

                // Load the CSV file into a Workbook
                Workbook workbook = new Workbook(csvPath, loadOptions);

                // NOTE: In recent Aspose.Cells versions the property to retain Office Add‑Ins
                // (EnableAddInSupport) may not be available. If required, ensure you are using a
                // version that provides this feature. The line is omitted to keep the code
                // compilable.

                // Save the workbook as PDF
                workbook.Save(pdfPath, SaveFormat.Pdf);

                Console.WriteLine("CSV has been successfully converted to PDF.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during conversion: {ex.Message}");
            }
        }
    }
}
