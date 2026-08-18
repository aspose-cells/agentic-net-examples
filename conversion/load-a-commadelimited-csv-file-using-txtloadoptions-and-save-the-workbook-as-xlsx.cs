// Title: C# – Convert a comma‑delimited CSV to XLSX using Aspose.Cells TxtLoadOptions
// Description: Demonstrates how to set a comma separator with TxtLoadOptions, load a CSV file into an Aspose.Cells Workbook, and save it as an XLSX workbook in C#.
// Keywords: Aspose.Cells CSV import C# | TxtLoadOptions separator | load CSV Aspose.Cells | save workbook as XLSX | C# CSV to Excel conversion | Aspose.Cells LoadOptions example | comma‑delimited CSV Aspose | Workbook SaveFormat.Xlsx | Aspose.Cells .NET CSV import
// Common Searches: Aspose.Cells load CSV with custom delimiter C# | How to convert CSV to XLSX using Aspose.Cells | TxtLoadOptions example for comma separated values | C# code to import CSV and export Excel with Aspose | SaveFormat.Xlsx usage in Aspose.Cells
// Developer Intent: Import a comma‑separated CSV via TxtLoadOptions and export it as an XLSX workbook.
// Use Cases: Transform user‑uploaded CSV reports into Excel files for analysis. | Automate batch conversion of multiple CSV files to XLSX in a data pipeline. | Generate Excel output from temporary CSV data created at runtime.
// AI Prompts: Generate C# code that reads a pipe‑delimited CSV with Aspose.Cells TxtLoadOptions and saves it as XLSX. | Explain how to configure TxtLoadOptions for different delimiters and optimize performance for large CSV files. | Provide a step‑by‑step tutorial for converting CSV to XLSX in a .NET Core console app using Aspose.Cells, including error handling.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvToXlsx
{
    // Demonstrates how to set a comma separator with TxtLoadOptions, load a CSV file into an Aspose.Cells Workbook, and save it as an XLSX workbook in C#.
    class Program
    {
        static void Main()
        {
            // Path to the source CSV file
            string csvPath = "input.csv";

            // Path for the resulting XLSX file
            string xlsxPath = "output.xlsx";

            // Create a sample CSV file (comma‑delimited)
            File.WriteAllText(csvPath, "Name,Age,Country\nJohn,30,USA\nAlice,25,UK\nBob,35,Canada");

            // ------------------------------------------------------------
            // Load the CSV file using TxtLoadOptions (comma separator)
            // ------------------------------------------------------------
            TxtLoadOptions loadOptions = new TxtLoadOptions();   // rule: TxtLoadOptions constructor
            loadOptions.Separator = ',';                         // set the CSV delimiter

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(csvPath, loadOptions); // rule: Workbook(string, LoadOptions)

            // ------------------------------------------------------------
            // Save the workbook as XLSX
            // ------------------------------------------------------------
            workbook.Save(xlsxPath, SaveFormat.Xlsx); // rule: Workbook.Save(string, SaveFormat)

            Console.WriteLine($"CSV file '{csvPath}' has been converted to XLSX file '{xlsxPath}'.");
        }
    }
}
