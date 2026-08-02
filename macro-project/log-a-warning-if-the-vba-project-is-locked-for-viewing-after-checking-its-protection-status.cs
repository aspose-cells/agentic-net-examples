// Title: Log a Warning When an Aspose.Cells VBA Project Is Locked for Viewing (C#)
// Description: C# example that creates a macro‑enabled workbook, accesses its VbaProject, reads the IsProtected and IsLockedForViewing flags, writes the status to the console, logs a warning if the project is locked for viewing, and saves the file as an XLSM workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | VbaProject | IsLockedForViewing | VBA protection | macro-enabled workbook | check VBA lock | log warning | Excel automation | .NET
// Common Searches: Aspose.Cells check if VBA project is locked for viewing | C# detect locked VBA project using Aspose.Cells | log warning when VBA project is locked in generated workbook | how to read VbaProject.IsLockedForViewing property | determine VBA protection status with Aspose.Cells .NET
// Developer Intent: Identify whether a workbook’s VBA project is locked for viewing and emit a warning when it is.
// Use Cases: Automated generation of macro‑enabled reports that must verify code visibility before distribution. | Pre‑deployment validation to ensure end users can view VBA code in delivered workbooks. | Security auditing of generated Excel files to flag hidden or protected VBA modules. | Conditional workflow that aborts or modifies saving logic when the VBA project is locked.
// AI Prompts: Provide C# code using Aspose.Cells that checks VbaProject.IsLockedForViewing and writes a warning to a log file with Serilog. | Create a reusable method that returns a boolean indicating a locked VBA project and logs via Microsoft.Extensions.Logging. | Explain how to handle exceptions when accessing VbaProject in a workbook that may not contain any macros, including fallback behavior.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCheck
{
    // C# example that creates a macro‑enabled workbook, accesses its VbaProject, reads the IsProtected and IsLockedForViewing flags, writes the status to the console, logs a warning if the project is locked for viewing, and saves the file as an XLSM workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (macro-enabled format will be used when saving)
            Workbook workbook = new Workbook();

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is protected
            bool isProtected = vbaProject.IsProtected;
            Console.WriteLine($"VBA Project Protected: {isProtected}");

            // Check if the VBA project is locked for viewing
            bool isLockedForViewing = vbaProject.IslockedForViewing;
            Console.WriteLine($"VBA Project Locked for Viewing: {isLockedForViewing}");

            // Log a warning if the project is locked for viewing
            if (isLockedForViewing)
            {
                Console.WriteLine("Warning: The VBA project is locked for viewing.");
            }

            // Example: protect and lock the VBA project (optional demonstration)
            // vbaProject.Protect(true, "myPassword");

            // Save the workbook as a macro-enabled file
            workbook.Save("VbaProjectCheck.xlsm", SaveFormat.Xlsm);
        }
    }
}
