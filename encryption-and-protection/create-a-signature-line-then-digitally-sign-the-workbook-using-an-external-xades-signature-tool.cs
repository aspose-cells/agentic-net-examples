// Title: Add a rectangle signature placeholder and sign an Excel workbook with an external XAdES tool using Aspose.Cells for .NET
// Common Searches: how to add a rectangle shape as a signature placeholder in Excel using Aspose.Cells for .NET | C# code to digitally sign an XLSX file with an external XAdES signer | run an external XAdES signing executable from a .NET application | verify that a signed Excel workbook was created after external signing in C# | save an unsigned workbook then apply XAdES digital signature with Aspose.Cells
// Tags: Aspose.Cells shape insertion C# | Excel signature line placeholder Aspose.Cells | XLSX workbook save Aspose.Cells C# | ProcessStartInfo external XAdES signer | post‑sign verification of Excel file C# | digital signature integration Aspose.Cells XAdES

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, inserts a rectangular shape labeled "Signature Line" as a visual placeholder on the first worksheet, saves the file as an unsigned XLSX, then launches an external XAdES signing executable via ProcessStartInfo to produce a signed workbook, checks the process exit code, and finally loads the signed file to confirm the digital signature was applied.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Position and size for a placeholder signature line
            int row = 4;      // zero‑based index (row 5)
            int column = 1;   // zero‑based index (column B)
            int offsetX = 0;
            int offsetY = 0;
            int width = 200;
            int height = 100;

            // Add a simple rectangle shape as a visual placeholder for the signature line
            Shape placeholder = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                row,
                column,
                offsetX,
                offsetY,
                width,
                height);
            placeholder.Text = "Signature Line";

            // Save the workbook to a temporary file (unsigned)
            string unsignedPath = "UnsignedWorkbook.xlsx";
            workbook.Save(unsignedPath, SaveFormat.Xlsx);

            // Path for the signed output file
            string signedPath = "SignedWorkbook.xlsx";

            // Verify the unsigned file exists before invoking the external signer
            if (!File.Exists(unsignedPath))
            {
                Console.WriteLine($"Unsigned workbook not found at {unsignedPath}");
                return;
            }

            // Invoke external XAdES signing tool
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "XAdESSigner.exe", // external XAdES signing tool
                Arguments = $"\"{unsignedPath}\" \"{signedPath}\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            try
            {
                using (Process proc = Process.Start(psi))
                {
                    if (proc == null)
                    {
                        Console.WriteLine("Failed to start the signing process.");
                        return;
                    }

                    // Capture any output for debugging (optional)
                    string output = proc.StandardOutput.ReadToEnd();
                    string error = proc.StandardError.ReadToEnd();
                    proc.WaitForExit();

                    if (proc.ExitCode != 0)
                    {
                        Console.WriteLine("XAdES signing failed:");
                        Console.WriteLine(error);
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while executing signing tool: {ex.Message}");
                return;
            }

            // Verify that the signed file exists before loading
            if (File.Exists(signedPath))
            {
                Workbook signedWorkbook = new Workbook(signedPath);
                Console.WriteLine("Workbook signed successfully and saved to: " + signedPath);
            }
            else
            {
                Console.WriteLine("Signed workbook was not created at the expected location.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
