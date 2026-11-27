// Title: Generate MHTML from an Aspose.Cells workbook with IE compatibility disabled and view it in Google Chrome (C#)
// Common Searches: c# save Aspose.Cells workbook as MHTML with IsIECompatible set to false | disable IE specific markup when exporting Excel to MHTML using Aspose.Cells | programmatically open an MHTML file in Google Chrome from C# | check MHTML output from Aspose.Cells for IE meta tags or CSS | export Excel images as Base64 in MHTML using Aspose.Cells HtmlSaveOptions
// Tags: Aspose.Cells HtmlSaveOptions MHTML export | IsIECompatible false Aspose.Cells HTML output | Base64 image embedding Aspose.Cells MHTML | C# launch Chrome to view MHTML file | detect IE specific markers in MHTML output

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The sample creates a Workbook, fills a 10x10 range, and saves it as an MHTML file using HtmlSaveOptions with IsIECompatible set to false and images exported as Base64. It then performs a basic check for IE‑specific CSS or meta tags and launches Google Chrome (if available) to display the generated MHTML for manual verification.
class Program
{
    static void Main()
    {
        // Path for the generated MHTML file
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleOutput.mht");

        try
        {
            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // 1. Create a workbook and populate it with sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill a 10x10 range with sample values
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // 2. Configure HTML (MHTML) save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml)
            {
                // Disable IE‑specific compatibility mode
                IsIECompatible = false,

                // Export images as Base64 to keep everything inside the MHTML package
                ExportImagesAsBase64 = true,

                // Optional: make the output more compact for testing
                ExportWorksheetCSSSeparately = false
            };

            // 3. Save the workbook as MHTML
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"MHTML file successfully saved to: {outputPath}");

            // 4. Verify that the file does not contain IE‑specific markers (basic check)
            if (File.Exists(outputPath))
            {
                string fileContent = File.ReadAllText(outputPath);
                bool containsIEFlag = fileContent.Contains("mso-") || fileContent.Contains("IE=Edge");
                Console.WriteLine("Basic verification of IE‑specific markers:");
                Console.WriteLine(containsIEFlag
                    ? "Warning: IE‑specific CSS or meta tags were found."
                    : "No IE‑specific markers detected – standard rendering expected.");
            }
            else
            {
                Console.WriteLine("Error: MHTML file was not created.");
            }

            // 5. Attempt to open the file in Google Chrome for visual confirmation
            // Adjust the Chrome executable path if necessary.
            string chromeExe = "chrome"; // assumes chrome is in the system PATH
            if (IsExecutableAvailable(chromeExe))
            {
                try
                {
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = chromeExe,
                        Arguments = $"\"{outputPath}\"",
                        UseShellExecute = true
                    };
                    Process.Start(psi);
                    Console.WriteLine("Chrome launched. Please verify the rendering manually.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Failed to launch Chrome:");
                    Console.WriteLine(ex.Message);
                }
            }
            else
            {
                Console.WriteLine("Google Chrome executable not found in PATH. Open the file manually in Chrome to verify rendering.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while generating or opening the MHTML file:");
            Console.WriteLine(ex.Message);
        }
    }

    private static bool IsExecutableAvailable(string exeName)
    {
        try
        {
            var paths = Environment.GetEnvironmentVariable("PATH")?.Split(Path.PathSeparator);
            if (paths == null) return false;

            foreach (var dir in paths)
            {
                string fullPath = Path.Combine(dir.Trim(), exeName + (OperatingSystem.IsWindows() ? ".exe" : ""));
                if (File.Exists(fullPath))
                    return true;
            }
        }
        catch
        {
            // Ignored – return false on any failure
        }
        return false;
    }
}
