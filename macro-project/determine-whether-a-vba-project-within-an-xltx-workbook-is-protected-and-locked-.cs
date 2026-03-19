using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Load the XLTX workbook (replace with the actual file path)
        Workbook workbook = new Workbook("template.xltx");

        // Get the VBA project from the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Check if a VBA project exists
        if (vbaProject != null)
        {
            // Output protection status
            Console.WriteLine("VBA Project IsProtected: " + vbaProject.IsProtected);
            // Output locked‑for‑viewing status
            Console.WriteLine("VBA Project IslockedForViewing: " + vbaProject.IslockedForViewing);
        }
        else
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
        }
    }
}