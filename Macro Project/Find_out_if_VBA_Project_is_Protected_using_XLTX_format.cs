using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Path to the XLTX template file
        string filePath = "template.xltx";

        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        // Load the workbook (XLTX format)
        Workbook workbook = new Workbook(filePath);

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        if (vbaProject != null)
        {
            // Output the protection status
            Console.WriteLine($"VBA Project Protected: {vbaProject.IsProtected}");

            // If the VbaProject class provides a lock status, display it
            // (Some versions expose IsLocked; if not, this line can be omitted)
            // Console.WriteLine($"VBA Project Locked: {vbaProject.IsLocked}");
        }
        else
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
        }
    }
}