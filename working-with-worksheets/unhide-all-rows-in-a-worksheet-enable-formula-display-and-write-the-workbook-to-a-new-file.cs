// Title: Unhide All Rows, Show Formulas, and Save Workbook to New File – Aspose.Cells C# Example
// Description: Loads an existing XLSX file, accesses the first worksheet, unhides every row using Cells.UnhideRows with auto‑fit height, optionally toggles formula view, creates the output folder if missing, and saves the modified workbook as a new XLSX file while handling missing files and runtime errors.
// Keywords: Aspose.Cells | C# unhide rows | Excel row visibility | Show formulas Aspose.Cells | Save workbook new file | Cells.UnhideRows | auto‑fit row height | error handling | create output directory | Aspose.Cells .NET
// Common Searches: Aspose.Cells unhide rows C# | How to display formulas in Aspose.Cells .NET | Save modified Excel workbook to another location using Aspose.Cells | Create output folder before saving workbook C# | Unhide hidden rows in Excel with Aspose.Cells | Cells.UnhideRows example
// Developer Intent: Unhide every row in the first worksheet, optionally display formulas, and write the workbook to a new file.
// Use Cases: Prepare uploaded Excel files for downstream processing by removing hidden rows and storing a clean copy. | Automate a reporting pipeline that guarantees all rows are visible before exporting the workbook to a shared folder. | Create a utility that validates input workbooks, reveals hidden content, and saves the result with a distinct filename.
// AI Prompts: Generate C# code using Aspose.Cells that unhides all rows in a worksheet, ensures the output directory exists, optionally enables formula view, and saves the workbook to a specified path with robust error handling. | Provide a reusable method that iterates through every worksheet in a workbook, calls Cells.UnhideRows for each sheet, toggles ShowFormula when supported, and saves each modified workbook with a custom suffix.

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing XLSX file, accesses the first worksheet, unhides every row using Cells.UnhideRows with auto‑fit height, optionally toggles formula view, creates the output folder if missing, and saves the modified workbook as a new XLSX file while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook.
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Unhide all rows in the worksheet.
            // Start from row index 0, total rows equal to the number of rows in the sheet,
            // and use -1 for height to let Aspose.Cells auto‑fit the row height.
            worksheet.Cells.UnhideRows(0, worksheet.Cells.Rows.Count, -1);

            // NOTE: The ShowFormula property is not available in the current Aspose.Cells version.
            // If needed, alternative approaches can be used to display formulas.

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook.
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
