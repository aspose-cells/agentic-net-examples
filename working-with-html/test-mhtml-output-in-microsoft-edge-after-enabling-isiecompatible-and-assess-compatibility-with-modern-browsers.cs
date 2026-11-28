// Title: Generate a single‑file MHTML workbook with Aspose.Cells in C# and preview it in Microsoft Edge using IsIECompatible mode
// Common Searches: how to save an Aspose.Cells workbook as a single MHTML file in C# | C# code to open a .mht file in Microsoft Edge after exporting from Aspose.Cells | validate MIME-Version and multipart/related headers of a generated MHTML file using .NET | export Excel workbook to MHTML with base64‑encoded images using Aspose.Cells | run Microsoft Edge from a C# application to display a generated .mht file
// Tags: Aspose.Cells HtmlSaveOptions MHTML export | IsIECompatible mode C# | single‑file MHTML workbook generation | base64 image embedding in MHTML | launch Microsoft Edge from .NET | MHTML MIME header validation .NET

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace MhtmlEdgeTest
{
    // The example creates a Workbook, fills a 10x10 range, configures HtmlSaveOptions with IsIECompatible=true, ExportImagesAsBase64=true, and ExportWorksheetCSSSeparately=false to produce a single MHTML file. It saves the file, verifies that the file exists and contains the required MIME-Version and multipart/related headers, and then launches Microsoft Edge (or the default handler) to display the generated .mht file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path for the generated MHTML file
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TestOutput.mht");

            try
            {
                // 1. Create a workbook and populate it with sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "SampleData";

                // Fill a 10x10 range with identifiable values
                for (int row = 0; row < 10; row++)
                {
                    for (int col = 0; col < 10; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // 2. Configure HTML save options for MHTML output
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.MHtml)
                {
                    // Enable IE‑compatible mode – this forces the output to be a single MHTML file
                    IsIECompatible = true,

                    // Export images as Base64 to keep everything inside the MHTML package
                    ExportImagesAsBase64 = true,

                    // Export the whole workbook as a single HTML (MHTML) file
                    ExportWorksheetCSSSeparately = false
                };

                // 3. Save the workbook as MHTML using the configured options
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"MHTML file successfully saved to: {outputPath}");

                // 4. Basic validation of the generated file
                ValidateMhtmlFile(outputPath);

                // 5. Launch Microsoft Edge to display the file
                LaunchEdge(outputPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred while generating or testing the MHTML file:");
                Console.Error.WriteLine(ex.ToString());
                Environment.Exit(1);
            }
        }

        private static void ValidateMhtmlFile(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                    throw new FileNotFoundException("MHTML file was not created.", filePath);

                FileInfo info = new FileInfo(filePath);
                if (info.Length == 0)
                    throw new InvalidDataException("MHTML file is empty.");

                // Read the first few kilobytes to look for MIME headers
                string header;
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                using (StreamReader reader = new StreamReader(fs))
                {
                    char[] buffer = new char[2048];
                    int read = reader.ReadBlock(buffer, 0, buffer.Length);
                    header = new string(buffer, 0, read);
                }

                if (!header.Contains("MIME-Version:") || !header.Contains("Content-Type: multipart/related"))
                    throw new InvalidDataException("MHTML file does not contain expected MIME headers.");

                Console.WriteLine("MHTML file validation passed.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Validation failed: {ex.Message}");
                throw;
            }
        }

        private static void LaunchEdge(string filePath)
        {
            try
            {
                // Edge executable locations (common on Windows 10/11)
                string[] possibleEdgePaths = new[]
                {
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"Microsoft\Edge\Application\msedge.exe"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"Microsoft\Edge\Application\msedge.exe")
                };

                string? edgeExe = null;
                foreach (var path in possibleEdgePaths)
                {
                    if (File.Exists(path))
                    {
                        edgeExe = path;
                        break;
                    }
                }

                ProcessStartInfo startInfo;
                if (!string.IsNullOrEmpty(edgeExe))
                {
                    startInfo = new ProcessStartInfo
                    {
                        FileName = edgeExe,
                        Arguments = $"\"{filePath}\"",
                        UseShellExecute = false
                    };
                    Console.WriteLine("Launching Microsoft Edge to display the MHTML file...");
                }
                else
                {
                    // Fallback to the default handler for .mht files
                    startInfo = new ProcessStartInfo
                    {
                        FileName = filePath,
                        UseShellExecute = true
                    };
                    Console.WriteLine("Microsoft Edge not found. Opening with the default system handler.");
                }

                Process? proc = Process.Start(startInfo);
                if (proc == null)
                    Console.WriteLine("Failed to start the viewer process. Please open the file manually.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error launching viewer: {ex.Message}");
            }
        }
    }
}
