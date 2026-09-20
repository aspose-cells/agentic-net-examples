// Title: Convert an Excel worksheet to a high‑resolution TIFF with Aspose.Cells and then to PDF using Ghostscript in C#
// AI Prompts: Generate C# code that loads an .xlsx file, uses Aspose.Cells to render the first worksheet as a 300 dpi TIFF, and saves the image to a file. | Write C# that starts the Ghostscript executable to convert the saved TIFF into a PDF, capturing standard output and error streams for logging.
// Common Searches: C# Aspose.Cells render worksheet to TIFF then convert to PDF with Ghostscript | How to invoke gswin64c from C# to transform a TIFF file into a PDF | Create high‑resolution image from Excel and programmatically convert it to PDF | Using Process.Start in .NET to run external PDF conversion tools
// Tags: Aspose.Cells TIFF rendering from Excel | Ghostscript TIFF to PDF conversion via C# | C# Process.Start external command execution | high‑resolution Excel worksheet image generation | command‑line PDF creation from image file

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example verifies the presence of an input.xlsx file, loads it with Aspose.Cells, renders the first worksheet to a 300 dpi TIFF using ImageOrPrintOptions, checks that the TIFF was created, then launches the Ghostscript command‑line tool (gswin64c) via a Process to convert the TIFF to output.pdf, handling stdout, stderr, and exit codes.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string tiffPath = "output.tiff";
            const string pdfPath = "output.pdf";

            // Verify the input Excel file exists
            if (!File.Exists(inputPath))
            {
                Console.Error.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure image rendering options (default format is PNG; Ghostscript can handle PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true,
                HorizontalResolution = 300,
                VerticalResolution = 300
                // ImageFormat property omitted to avoid compatibility issues
            };

            // Render the first worksheet to an image file
            SheetRender sheetRender = new SheetRender(workbook.Worksheets[0], imgOptions);
            sheetRender.ToImage(0, tiffPath);

            // Ensure the image file was created
            if (!File.Exists(tiffPath))
            {
                Console.Error.WriteLine($"Failed to create image file: {tiffPath}");
                return;
            }

            // Prepare Ghostscript conversion command
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "gswin64c", // Ghostscript executable
                Arguments = $"-dNOPAUSE -dBATCH -sDEVICE=pdfwrite -sOutputFile={pdfPath} {tiffPath}",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            // Execute the conversion process
            using (Process? proc = Process.Start(startInfo))
            {
                if (proc == null)
                {
                    Console.Error.WriteLine("Failed to start the conversion process.");
                    return;
                }

                proc.WaitForExit();

                string output = proc.StandardOutput.ReadToEnd();
                string error = proc.StandardError.ReadToEnd();

                if (proc.ExitCode != 0)
                {
                    Console.Error.WriteLine($"Conversion failed (exit code {proc.ExitCode}): {error}");
                }
                else
                {
                    Console.WriteLine("Conversion succeeded.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
