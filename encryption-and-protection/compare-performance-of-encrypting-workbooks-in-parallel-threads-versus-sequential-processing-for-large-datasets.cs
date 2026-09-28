// Title: Performance comparison of sequential vs parallel password encryption for large Excel workbooks using Aspose.Cells in C#
// AI Prompts: Generate C# code that creates multiple 10,000‑row by 50‑column Excel workbooks, applies a password with Aspose.Cells, saves them, and prints the elapsed time for both a for‑loop and a Parallel.For implementation. | Modify the parallel encryption sample to cap the degree of parallelism to the number of logical processors and calculate the percentage speed‑up over the sequential run. | Add diagnostics to the benchmark that log CPU and memory usage while each workbook is being encrypted in parallel with Aspose.Cells.
// Common Searches: how fast is Aspose.Cells password protection when using Parallel.For in C# | benchmark sequential versus multi‑threaded Excel file encryption Aspose.Cells .NET | measure encryption time for large Excel workbooks with Aspose.Cells C# example | optimal thread count for encrypting many Excel files with Aspose.Cells
// Tags: parallel workbook encryption Aspose.Cells .NET | sequential Excel password protection timing | Aspose.Cells encryption performance benchmark | multi‑threaded Excel file saving with password | large dataset workbook encryption C#

using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// The sample creates ten Excel workbooks (10,000 rows × 50 columns), encrypts each with a password using Aspose.Cells, and measures the elapsed time for a sequential for‑loop and a Parallel.For approach, deleting files between runs to ensure a fair performance comparison.
class Program
{
    // Number of workbooks to create and encrypt
    const int WorkbookCount = 10;

    // Size of each workbook (rows x columns)
    const int RowCount = 10000;
    const int ColumnCount = 50;

    // Encryption password
    const string Password = "Secret123";

    // Folder to store generated files
    static readonly string OutputFolder = Path.Combine(Directory.GetCurrentDirectory(), "EncryptedWorkbooks");

    static void Main()
    {
        try
        {
            // Ensure output directory exists
            Directory.CreateDirectory(OutputFolder);

            // Prepare file paths
            string[] filePaths = new string[WorkbookCount];
            for (int i = 0; i < WorkbookCount; i++)
            {
                filePaths[i] = Path.Combine(OutputFolder, $"Workbook_{i + 1}.xlsx");
            }

            // ---------- Sequential Encryption ----------
            Stopwatch swSequential = Stopwatch.StartNew();

            for (int i = 0; i < WorkbookCount; i++)
            {
                try
                {
                    // Create a workbook with large data
                    Workbook wb = CreateLargeWorkbook();

                    // Set password for opening the workbook
                    wb.Settings.Password = Password;

                    // Save encrypted workbook
                    wb.Save(filePaths[i]);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing workbook {i + 1} sequentially: {ex.Message}");
                }
            }

            swSequential.Stop();
            Console.WriteLine($"Sequential encryption time: {swSequential.ElapsedMilliseconds} ms");

            // ---------- Parallel Encryption ----------
            // Delete previously saved files to avoid confusion
            foreach (var path in filePaths)
            {
                try
                {
                    if (File.Exists(path))
                        File.Delete(path);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error deleting file '{path}': {ex.Message}");
                }
            }

            Stopwatch swParallel = Stopwatch.StartNew();

            Parallel.For(0, WorkbookCount, i =>
            {
                try
                {
                    // Each thread works with its own workbook instance
                    Workbook wb = CreateLargeWorkbook();

                    // Set password for opening the workbook
                    wb.Settings.Password = Password;

                    // Save encrypted workbook
                    wb.Save(filePaths[i]);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing workbook {i + 1} in parallel: {ex.Message}");
                }
            });

            swParallel.Stop();
            Console.WriteLine($"Parallel encryption time: {swParallel.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    // Helper method to create a workbook filled with dummy data
    static Workbook CreateLargeWorkbook()
    {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate cells with sample data
        for (int row = 0; row < RowCount; row++)
        {
            for (int col = 0; col < ColumnCount; col++)
            {
                cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        return wb;
    }
}
