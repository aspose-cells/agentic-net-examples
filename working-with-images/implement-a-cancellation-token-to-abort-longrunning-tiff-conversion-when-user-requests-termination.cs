// Title: Implement CancellationToken to abort Excel-to-TIFF conversion with Aspose.Cells in C#
// Common Searches: how to cancel Aspose.Cells workbook.Save to TIFF in a .NET application | C# example for using CancellationToken with Aspose.Cells image export | abort long running Excel to TIFF conversion using Aspose.Cells and CancellationToken | remove partially created TIFF file when conversion is cancelled in Aspose.Cells | using Task.Run and CancellationToken to stop Aspose.Cells TIFF export
// Tags: Aspose.Cells cancellation token for TIFF export | C# workbook.Save cancellation handling | Excel to TIFF conversion abort | partial TIFF file cleanup Aspose.Cells | asynchronous image export with Aspose.Cells

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The ConvertToTiff method validates file paths, checks the supplied CancellationToken before and after loading the workbook, runs workbook.Save to TIFF on a background task, waits with the token, and on cancellation deletes any partially written TIFF file while ensuring the workbook is disposed.
    public class TiffConverter
    {
        /// <param name="inputPath">Full path to the source Excel file.</param>
        /// <param name="outputPath">Full path where the TIFF file will be saved.</param>
        /// <param name="cancellationToken">Token used to request cancellation.</param>
        public void ConvertToTiff(string inputPath, string outputPath, CancellationToken cancellationToken)
        {
            // Validate arguments
            if (string.IsNullOrEmpty(inputPath))
                throw new ArgumentException("Input path must be provided.", nameof(inputPath));
            if (string.IsNullOrEmpty(outputPath))
                throw new ArgumentException("Output path must be provided.", nameof(outputPath));
            if (!File.Exists(inputPath))
                throw new FileNotFoundException("Input file not found.", inputPath);

            Workbook workbook = null;
            try
            {
                // Check for cancellation before loading
                cancellationToken.ThrowIfCancellationRequested();

                // Load workbook (potentially long‑running)
                workbook = new Workbook(inputPath);

                // Check again after loading
                cancellationToken.ThrowIfCancellationRequested();

                // Save as TIFF on a separate task to observe cancellation
                var saveTask = Task.Run(() =>
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    workbook.Save(outputPath, SaveFormat.Tiff);
                }, cancellationToken);

                // Wait for completion, propagating cancellation if requested
                saveTask.Wait(cancellationToken);
            }
            catch (OperationCanceledException)
            {
                // Clean up partially written file on cancellation
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
                throw;
            }
            finally
            {
                // Ensure resources are released
                workbook?.Dispose();
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            // Input and output paths can be supplied via command‑line arguments
            string inputPath = args.Length > 0 ? args[0] : "sample.xlsx";
            string outputPath = args.Length > 1 ? args[1] : "output.tiff";

            // Verify the input file exists before proceeding
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            var converter = new TiffConverter();
            using var cts = new CancellationTokenSource();

            try
            {
                converter.ConvertToTiff(inputPath, outputPath, cts.Token);
                Console.WriteLine($"TIFF file saved to: {outputPath}");
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("Conversion was cancelled.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
