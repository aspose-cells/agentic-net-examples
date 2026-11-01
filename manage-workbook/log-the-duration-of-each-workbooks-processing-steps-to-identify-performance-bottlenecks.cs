// Title: Measure and log each processing step (load, calculate formulas, auto‑fit columns, save) of an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to benchmark workbook loading time with Aspose.Cells in C# | C# code to time formula calculation in an Aspose.Cells workbook | measure auto‑fit column performance using Aspose.Cells .NET | log Excel file save duration with Aspose.Cells library | profile Aspose.Cells workbook processing steps in a .NET application
// Tags: Aspose.Cells workbook load timing | Aspose.Cells formula calculation performance | Aspose.Cells auto‑fit columns duration | Aspose.Cells workbook save latency | C# Stopwatch performance logging for Excel files

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to use Aspose.Cells in C# to load an Excel workbook, calculate its formulas, auto‑fit columns, and save the result while measuring each operation with Stopwatch. It logs the elapsed milliseconds for loading, formula calculation, column auto‑fit, and saving, and includes error handling for missing files, password‑protected workbooks, and I/O issues.
class WorkbookProcessor
{
    // Process a workbook and log the duration of each step
    public static void ProcessWorkbook(string inputPath, string outputPath)
    {
        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found – \"{inputPath}\"");
            return;
        }

        Workbook workbook = null;

        // Load workbook with error handling (e.g., password‑protected files)
        try
        {
            Stopwatch swLoad = Stopwatch.StartNew();
            // LOAD_WORKBOOK rule
            workbook = new Workbook(inputPath);
            swLoad.Stop();
            Console.WriteLine($"Load time: {swLoad.ElapsedMilliseconds} ms");
        }
        catch (CellsException ex)
        {
            // Aspose.Cells may throw a generic CellsException for password‑protected files
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Error: The workbook is password protected. Provide a valid password.");
            }
            else
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Ensure workbook was loaded before proceeding
        if (workbook == null)
        {
            Console.WriteLine("Error: Workbook could not be loaded.");
            return;
        }

        // Calculate formulas
        try
        {
            Stopwatch swCalc = Stopwatch.StartNew();
            workbook.CalculateFormula();
            swCalc.Stop();
            Console.WriteLine($"Formula calculation time: {swCalc.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during formula calculation: {ex.Message}");
        }

        // Auto-fit columns (example of an additional processing step)
        try
        {
            Stopwatch swAutoFit = Stopwatch.StartNew();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.AutoFitColumns();
            swAutoFit.Stop();
            Console.WriteLine($"Auto-fit columns time: {swAutoFit.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during auto‑fit columns: {ex.Message}");
        }

        // Ensure output directory exists
        try
        {
            string outDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
            {
                Directory.CreateDirectory(outDir);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error preparing output directory: {ex.Message}");
            return;
        }

        // Save workbook
        try
        {
            Stopwatch swSave = Stopwatch.StartNew();
            // SAVE_WORKBOOK rule
            workbook.Save(outputPath);
            swSave.Stop();
            Console.WriteLine($"Save time: {swSave.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }

    static void Main(string[] args)
    {
        // Example usage
        string inputFile = "input.xlsx";
        string outputFile = "output.xlsx";

        // Process the workbook and log durations
        ProcessWorkbook(inputFile, outputFile);
    }
}
