using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Path to the macro-enabled Excel file
        string filePath = "sample.xlsm";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(filePath);

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Retrieve the protection status of the VBA project
        bool isProtected = vbaProject.IsProtected;

        // Output the result in TSV format: FilePath<TAB>IsProtected
        Console.WriteLine($"{filePath}\t{isProtected}");
    }
}