using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProtection
{
    static void Main()
    {
        // Path to the SpreadsheetML (Excel 2003 XML) file
        string inputPath = "sample.xml";

        // Load the workbook from the SpreadsheetML file
        Workbook workbook = new Workbook(inputPath);

        // Verify that the workbook contains a VBA project
        if (workbook.HasMacro && workbook.VbaProject != null)
        {
            // Determine whether the VBA project is protected
            bool isProtected = workbook.VbaProject.IsProtected;
            Console.WriteLine($"VBA Project Protected: {isProtected}");
        }
        else
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
        }
    }
}