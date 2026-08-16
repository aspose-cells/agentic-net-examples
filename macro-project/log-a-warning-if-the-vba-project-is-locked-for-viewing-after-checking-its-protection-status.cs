// Title: Log a Warning When a VBA Project Is Locked for Viewing – Aspose.Cells for .NET
// Description: Demonstrates how to create a macro‑enabled workbook, protect its VBA project with a password, check the VbaProject.IsProtected flag, and write a warning to the console if the project is locked for viewing.
// Keywords: Aspose.Cells | VBA project protection | .NET | C# | VbaProject.IsProtected | locked for viewing | log warning | macro-enabled workbook | Excel VBA security
// Common Searches: Aspose.Cells check if VBA project is protected | log warning when VBA project is locked | C# detect locked VBA project in .xlsm | VbaProject.IsProtected example | protect VBA project with Aspose.Cells
// Developer Intent: Identify whether a workbook's VBA project is password‑protected and emit a warning when it is.
// Use Cases: Validate VBA protection status after programmatically securing a macro‑enabled workbook. | Scan a collection of .xlsm files and flag any that have a locked VBA project before deployment. | Integrate a protection‑check step into CI/CD pipelines to prevent accidental VBA locking.
// AI Prompts: Generate C# code using Aspose.Cells that opens an .xlsm file, checks VbaProject.IsProtected, and logs a warning if true. | Create a reusable method that receives a Workbook object, verifies VBA project protection, and writes a warning via a logging framework. | Show how to handle exceptions while protecting a VBA project and then confirming its locked status with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a macro‑enabled workbook, protect its VBA project with a password, check the VbaProject.IsProtected flag, and write a warning to the console if the project is locked for viewing.
    public class VbaProjectLockWarningDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook wb = new Workbook();

                // Save as a macro‑enabled workbook to ensure a VBA project exists
                string tempPath = "temp.xlsm";
                wb.Save(tempPath, SaveFormat.Xlsm);

                // Load the workbook only if the temporary file exists
                if (File.Exists(tempPath))
                {
                    wb = new Workbook(tempPath);
                    try
                    {
                        File.Delete(tempPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not delete temporary file. {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Temporary file '{tempPath}' not found.");
                    return;
                }

                // Protect the VBA project with a password
                try
                {
                    wb.VbaProject.Protect(true, "secret");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error protecting VBA project: {ex.Message}");
                }

                // Check protection status
                bool isProtected = wb.VbaProject.IsProtected;

                Console.WriteLine($"VBA Project Protected: {isProtected}");

                // Log a warning if the VBA project is protected (locked for viewing)
                if (isProtected)
                {
                    Console.WriteLine("Warning: The VBA project is locked for viewing.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            VbaProjectLockWarningDemo.Run();
        }
    }
}
