// Title: Benchmark opening time of an encrypted Excel workbook using Aspose.Cells for .NET
// Description: This C# console example checks for a password‑protected .xlsx file, sets the password in LoadOptions, measures the load duration with Stopwatch, prints the elapsed milliseconds and the workbook’s IsEncrypted flag, and then disposes the Workbook object.
// Keywords: Aspose.Cells | C# encrypted workbook | Excel password protection | load performance | benchmark workbook opening | Stopwatch timing | LoadOptions password | measure decryption time | Aspose.Cells performance testing | encrypted .xlsx loading .NET
// Common Searches: How to benchmark encrypted Excel file load time with Aspose.Cells | Measure opening latency of password protected workbook in C# | Aspose.Cells performance test for encrypted .xlsx | Timing decryption of Excel using Aspose.Cells .NET | Log load time of protected workbook Aspose.Cells
// Developer Intent: Determine the duration required for Aspose.Cells to open a password‑protected Excel file and capture that metric.
// Use Cases: Assess decryption overhead for large encrypted workbooks in a high‑throughput application | Integrate load‑time logging into monitoring dashboards for secure Excel processing | Compare opening speeds of workbooks encrypted with different algorithms or password complexities | Validate that encryption does not exceed SLA latency requirements
// AI Prompts: Generate C# code that records the time to open an encrypted Excel workbook with Aspose.Cells and writes the result to a log file. | Show how to run the opening test multiple times and compute average, min, and max load times. | Explain how to export the timing data to CSV or JSON for further analysis. | Provide guidance on handling CellsException for incorrect passwords while still measuring elapsed time.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPerformanceDemo
{
    // This C# console example checks for a password‑protected .xlsx file, sets the password in LoadOptions, measures the load duration with Stopwatch, prints the elapsed milliseconds and the workbook’s IsEncrypted flag, and then disposes the Workbook object.
    class Program
    {
        static void Main()
        {
            // Path to the encrypted workbook file
            string filePath = "encrypted.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File \"{filePath}\" not found.");
                return;
            }

            // Password required to open the workbook
            string password = "password";

            // Prepare load options with the password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };

            Workbook workbook = null;
            Stopwatch sw = Stopwatch.StartNew();

            try
            {
                // Load the encrypted workbook
                workbook = new Workbook(filePath, loadOptions);
            }
            catch (CellsException ex)
            {
                // Handle invalid password or other loading issues
                Console.WriteLine($"Failed to open workbook: {ex.Message}");
                return;
            }
            finally
            {
                sw.Stop();
            }

            // Log performance metrics
            Console.WriteLine($"Time taken to open encrypted workbook: {sw.ElapsedMilliseconds} ms");

            // Verify that the workbook reports it is encrypted
            Console.WriteLine($"Workbook.IsEncrypted: {workbook.Settings.IsEncrypted}");

            // Dispose the workbook when done
            workbook.Dispose();
        }
    }
}
