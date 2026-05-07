using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProtection
{
    static void Main()
    {
        // Path to the MHT file containing the workbook
        string filePath = Path.Combine(Environment.CurrentDirectory, "sample.mht");

        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        // Load the MHT file into a Workbook object
        Workbook workbook = new Workbook(filePath);

        // Verify that the workbook actually contains a VBA project
        if (workbook.HasMacro && workbook.VbaProject != null)
        {
            // Retrieve the protection status of the VBA project
            bool isProtected = workbook.VbaProject.IsProtected;

            // Output the result
            Console.WriteLine("VBA project is protected: " + isProtected);
        }
        else
        {
            Console.WriteLine("The file does not contain a VBA project.");
        }
    }
}