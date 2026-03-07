using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ImageConversionMonitor
{
    static void Main()
    {
        // Paths for source workbook and output image
        string sourcePath = "input.xlsx";
        string outputImage = "output_page_0.png";

        // Create a time‑based interrupt monitor (no exception on timeout)
        SystemTimeInterruptMonitor monitor = new SystemTimeInterruptMonitor(false);

        // Attach the monitor to LoadOptions so it can interrupt loading if needed
        LoadOptions loadOptions = new LoadOptions
        {
            InterruptMonitor = monitor
        };

        // Start monitoring before loading the workbook (e.g., 2 seconds limit)
        monitor.StartMonitor(2000);

        Workbook workbook;
        try
        {
            // Load the workbook with the interrupt monitor attached
            workbook = new Workbook(sourcePath, loadOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Loading interrupted or failed: " + ex.Message);
            return;
        }

        // Reset the monitor for the rendering operation (e.g., 3 seconds limit)
        monitor.StartMonitor(3000);

        // Configure image rendering options
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = Aspose.Cells.Drawing.ImageType.Png,
            PrintWithStatusDialog = true   // Show status dialog during conversion
        };

        // Create a renderer for the workbook
        WorkbookRender renderer = new WorkbookRender(workbook, options);

        try
        {
            // Render the first page of the workbook to an image file
            renderer.ToImage(0, outputImage);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Rendering interrupted or failed: " + ex.Message);
            return;
        }

        // After rendering, check whether the operation was interrupted
        if (monitor.IsInterruptionRequested)
        {
            Console.WriteLine("Conversion was interrupted due to time limit.");
        }
        else
        {
            Console.WriteLine($"Image conversion completed successfully: {outputImage}");
        }
    }
}