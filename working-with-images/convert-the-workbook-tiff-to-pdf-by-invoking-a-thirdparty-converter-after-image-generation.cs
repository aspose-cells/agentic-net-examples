using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsTiffToPdf
{
    class Program
    {
        static void Main()
        {
            // Paths for source Excel, intermediate TIFF and final PDF
            string excelPath = "input.xlsx";
            string tiffPath = "intermediate.tiff";
            string pdfPath = "output.pdf";

            // Load the workbook (create/load lifecycle handled by Aspose.Cells)
            Workbook workbook = new Workbook(excelPath);

            // Configure rendering options for TIFF output
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Tiff,
                TiffCompression = TiffCompression.CompressionLZW,
                HorizontalResolution = 300,
                VerticalResolution = 300
            };

            // Render the entire workbook to a multi‑page TIFF file
            WorkbookRender renderer = new WorkbookRender(workbook, renderOptions);
            renderer.ToImage(tiffPath); // Uses the ToImage(string) method to save as TIFF

            // ------------------------------------------------------------
            // Invoke a third‑party converter to transform the TIFF to PDF.
            // This example uses a command‑line tool (e.g., ImageMagick's convert)
            // Replace the command and arguments with the actual converter you have.
            // ------------------------------------------------------------
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "magick", // Example: ImageMagick's 'magick' command
                    Arguments = $"\"{tiffPath}\" \"{pdfPath}\"",
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using (Process proc = Process.Start(psi))
                {
                    proc.WaitForExit();

                    string output = proc.StandardOutput.ReadToEnd();
                    string error = proc.StandardError.ReadToEnd();

                    if (proc.ExitCode == 0)
                    {
                        Console.WriteLine($"Successfully converted TIFF to PDF: {pdfPath}");
                    }
                    else
                    {
                        Console.WriteLine($"Conversion failed with exit code {proc.ExitCode}");
                        Console.WriteLine($"Error: {error}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception while invoking converter: {ex.Message}");
            }

            // Optional: clean up the intermediate TIFF file
            if (File.Exists(tiffPath))
            {
                File.Delete(tiffPath);
            }
        }
    }
}