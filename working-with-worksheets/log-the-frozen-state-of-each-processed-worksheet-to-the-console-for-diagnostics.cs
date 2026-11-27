// Title: How to log freeze pane settings of each worksheet to the console using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells read frozen pane coordinates from all worksheets | Aspose.Cells get PaneState and frozen rows/columns for each sheet | log worksheet freeze pane information to console in .NET | diagnose frozen panes in an Excel workbook using Aspose.Cells API | retrieve freeze pane details with Worksheet.GetFreezedPanes in C#
// Tags: Aspose.Cells Worksheet.GetFreezedPanes usage | Aspose.Cells PaneState inspection | log worksheet freeze pane details C# | diagnostic freeze panes Aspose.Cells | Excel workbook freeze pane diagnostics .NET

using System;
using System.IO;
using Aspose.Cells;

namespace FreezePaneDiagnostics
{
    // The example loads an Excel workbook (from a file path or a generated sample), iterates through every worksheet, and writes to the console the worksheet name, PaneState, presence of frozen panes, and detailed freeze coordinates (row, column, frozen rows, frozen columns) using Worksheet.PaneState and Worksheet.GetFreezedPanes. An optional output path can be provided to save the workbook after diagnostics.
    class Program
    {
        /// <param name="args">
        /// Optional command‑line arguments.
        /// args[0] – path to an existing Excel file to process.
        /// args[1] – (optional) path where the workbook will be saved after processing.
        /// </param>
        static void Main(string[] args)
        {
            try
            {
                Workbook workbook;

                if (args.Length > 0 && File.Exists(args[0]))
                {
                    // Load workbook from the supplied file path.
                    Console.WriteLine($"Loading workbook from \"{args[0]}\"...");
                    workbook = new Workbook(args[0]);
                }
                else
                {
                    // No valid input file supplied – create a sample workbook.
                    Console.WriteLine("No valid input file supplied. Creating a sample workbook for demonstration.");
                    workbook = CreateSampleWorkbook();
                }

                // Iterate through all worksheets and log their frozen state.
                Console.WriteLine("\n--- Freeze Pane Diagnostics ---");
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    LogFreezeState(sheet);
                }

                // If an output path is provided, save the workbook.
                if (args.Length > 1)
                {
                    string outputPath = args[1];
                    Console.WriteLine($"\nSaving workbook to \"{outputPath}\"...");
                    workbook.Save(outputPath);
                }

                Console.WriteLine("\nProcessing completed successfully.");
            }
            catch (Exception ex)
            {
                // Do not hide failures – report the full exception.
                Console.Error.WriteLine($"ERROR: {ex.GetType().Name} - {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
                Environment.Exit(1);
            }
        }

        /// <summary>
        /// Logs the frozen pane information for a single worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet to inspect.</param>
        private static void LogFreezeState(Worksheet worksheet)
        {
            // Retrieve the high‑level pane state (enum) and the detailed freeze information.
            PaneStateType paneState = worksheet.PaneState;
            bool hasFreeze = worksheet.GetFreezedPanes(
                out int row,
                out int column,
                out int frozenRows,
                out int frozenColumns);

            Console.WriteLine($"Worksheet \"{worksheet.Name}\":");
            Console.WriteLine($"  PaneState          : {paneState}");
            Console.WriteLine($"  HasFreezePanes     : {hasFreeze}");

            if (hasFreeze)
            {
                Console.WriteLine($"  Freeze Position    : Row={row}, Column={column}");
                Console.WriteLine($"  Frozen Rows Count  : {frozenRows}");
                Console.WriteLine($"  Frozen Columns Count: {frozenColumns}");
            }
        }

        /// <summary>
        /// Creates a sample workbook with two worksheets:
        ///   - Sheet1 has frozen panes (first row and first column).
        ///   - Sheet2 has no frozen panes.
        /// </summary>
        /// <returns>A populated Workbook instance.</returns>
        private static Workbook CreateSampleWorkbook()
        {
            Workbook wb = new Workbook();

            // Worksheet 1 – with frozen panes.
            Worksheet ws1 = wb.Worksheets[0];
            ws1.Name = "FrozenDemo";
            ws1.Cells["A1"].PutValue("Header");
            ws1.Cells["A2"].PutValue("Data 1");
            ws1.Cells["B2"].PutValue("Data 2");
            // Freeze first row and first column (cell B2 is the top‑left visible cell after freeze).
            ws1.FreezePanes(1, 1, 1, 1);

            // Worksheet 2 – without frozen panes.
            Worksheet ws2 = wb.Worksheets.Add("NoFreeze");
            ws2.Cells["A1"].PutValue("Just some data");
            ws2.Cells["B1"].PutValue(123);
            ws2.Cells["C1"].PutValue(DateTime.Now);

            return wb;
        }
    }
}
