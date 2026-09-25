// Title: Implement a CancellationToken to abort a long‑running Aspose.Cells Excel‑to‑TIFF conversion in C#
// AI Prompts: Add a CancellationToken argument to the async ConvertToTiffAsync method and invoke token.ThrowIfCancellationRequested before calling workbook.Save. | In Main, create a CancellationTokenSource, start a helper task that cancels after a user‑defined timeout or input, and wrap the conversion call in a try‑catch for OperationCanceledException. | Execute the workbook.Save operation inside Task.Run, passing the cancellation token so the save can be interrupted safely.
// Common Searches: how to use CancellationToken with Aspose.Cells SaveFormat.Tiff in .NET | cancel long running Excel to TIFF conversion using Aspose.Cells | abort Aspose.Cells workbook.Save when user presses cancel button | async Excel to TIFF export with timeout and cancellation token C# | handle OperationCanceledException during Aspose.Cells image export
// Tags: cancellation support Aspose.Cells TIFF export | async workbook.Save with token .NET | user‑initiated abort of long running conversion | background task cancellation for Excel to image | operationcanceledexception handling Aspose.Cells

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;

// The sample loads an Excel workbook with Aspose.Cells, converts it to a TIFF image on a background thread, and uses a CancellationTokenSource that can be triggered by user input or a timeout. The conversion respects the token, allowing the operation to be cancelled and handling OperationCanceledException along with other possible errors.
class Program
{
    static async Task Main(string[] args)
    {
        // Paths for source workbook and target TIFF file
        string inputPath = "input.xlsx";
        string outputPath = "output.tiff";

        // Ensure the source file exists before proceeding
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Source file not found: {inputPath}");
            return;
        }

        // Create a cancellation token source that can be triggered by the user
        using var cts = new CancellationTokenSource();

        // Example: automatically cancel after 5 seconds (replace with real user input handling)
        _ = Task.Run(() =>
        {
            Thread.Sleep(5000);
            cts.Cancel(); // user requested termination
        });

        try
        {
            // Perform the conversion respecting the cancellation token
            await ConvertToTiffAsync(inputPath, outputPath, cts.Token);
            Console.WriteLine("TIFF conversion completed successfully.");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("TIFF conversion was cancelled by the user.");
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"File not found: {ex.FileName}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    static Task ConvertToTiffAsync(string sourceFile, string targetFile, CancellationToken token)
    {
        // Run the conversion on a background thread so it can be cancelled
        return Task.Run(() =>
        {
            // Verify source file exists
            if (!File.Exists(sourceFile))
                throw new FileNotFoundException("Source workbook not found.", sourceFile);

            // Load the workbook (lifecycle rule: load)
            var workbook = new Workbook(sourceFile);

            // Check for cancellation before the potentially long save operation
            token.ThrowIfCancellationRequested();

            // Save the workbook as TIFF (lifecycle rule: save)
            workbook.Save(targetFile, SaveFormat.Tiff);
        }, token);
    }
}
