// Title: C# Aspose.Cells: Detect Password‑Protected VBA Project in an .xlsm Workbook
// Description: Load an .xlsm file with Aspose.Cells, access its VbaProject, and use the IsProtected property to determine if the VBA macro project is secured with a password. The example prints a Boolean result.
// Keywords: Aspose.Cells VBA protection | C# VbaProject.IsProtected | check macro password .NET | detect password‑protected VBA project | load .xlsm workbook Aspose | Excel macro security C#
// Common Searches: how to know if a VBA project is password protected using Aspose.Cells | C# example for VbaProject.IsProtected | Aspose.Cells detect macro password in .xlsm | check VBA project protection .NET | Aspose.Cells VBA security status
// Developer Intent: Determine whether the VBA project embedded in a loaded workbook is locked with a password.
// Use Cases: Skip or warn users when processing workbooks that contain password‑protected macros. | Log VBA protection status across multiple files for compliance reporting. | Conditionally extract or modify macros only when the project is not secured.
// AI Prompts: Create a C# method that receives an .xlsm path, returns true if its VBA project is password protected, and safely handles files without a VBA project. | Generate code that opens an Excel workbook, checks VbaProject.IsProtected, logs the outcome, and proceeds to read macros only when unprotected. | Write a script that scans a folder of .xlsm files, reports each file's VBA protection state, and produces a summary CSV.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Load an .xlsm file with Aspose.Cells, access its VbaProject, and use the IsProtected property to determine if the VBA macro project is secured with a password. The example prints a Boolean result.
class Program
{
    static void Main()
    {
        // Load the workbook that may contain a VBA project
        Workbook workbook = new Workbook("input.xlsm");

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Determine whether the VBA project is protected with a password
        bool isPasswordProtected = vbaProject.IsProtected;

        // Output the result
        Console.WriteLine("VBA Project password protected: " + isPasswordProtected);
    }
}
