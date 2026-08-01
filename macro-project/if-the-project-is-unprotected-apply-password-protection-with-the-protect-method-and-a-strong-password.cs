// Title: Protect a VBA Project with a Password in an Excel .xlsm Workbook using Aspose.Cells for .NET (C#)
// Description: Shows how to create or load a workbook, detect an unprotected VBA project, apply a strong password with Workbook.VbaProject.Protect (without locking the code for viewing), and save the file as a macro‑enabled .xlsm using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | VBA project protection | Workbook.VbaProject.Protect | macro-enabled workbook | password protection | .xlsm | .NET Excel automation | secure VBA code | protect VBA programmatically
// Common Searches: Aspose.Cells protect VBA project C# | How to add password to VBA project with Aspose.Cells | Check if VBA project is protected before applying password .NET | Save macro-enabled workbook after VBA protection | C# code to secure VBA in .xlsm using Aspose
// Developer Intent: Add password protection to an unprotected VBA project in an Excel workbook via Aspose.Cells for .NET.
// Use Cases: Automatically secure newly generated .xlsm files before distribution. | Batch‑process a folder of macro‑enabled workbooks to enforce VBA password policies. | Integrate VBA project protection into CI/CD pipelines for Excel add‑ins built with Aspose.Cells.
// AI Prompts: Generate C# code that uses Aspose.Cells to protect a VBA project only when it is not already secured, accepting the password as an argument. | Write a script that scans a directory for .xlsm files and applies a strong password to each workbook's VBA project using Aspose.Cells. | Create a reusable Aspose.Cells utility class that checks VbaProject.IsProtected and applies Workbook.VbaProject.Protect with a configurable password.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Shows how to create or load a workbook, detect an unprotected VBA project, apply a strong password with Workbook.VbaProject.Protect (without locking the code for viewing), and save the file as a macro‑enabled .xlsm using Aspose.Cells for .NET.
class ProtectVbaProject
{
    static void Main()
    {
        // Create a new workbook (or load an existing one with a VBA project)
        Workbook workbook = new Workbook();

        // Check if the VBA project is already protected
        if (!workbook.VbaProject.IsProtected)
        {
            // Protect the VBA project (not locked for viewing) with a strong password
            workbook.VbaProject.Protect(false, "Str0ngP@ssw0rd!2026");
        }

        // Save the workbook as a macro‑enabled file
        workbook.Save("ProtectedVbaProject.xlsm", SaveFormat.Xlsm);
    }
}
