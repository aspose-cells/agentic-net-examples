using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the MHTML file
        string filePath = "sample.mht";

        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {Path.GetFullPath(filePath)}");
            return;
        }

        // Load the workbook from the MHTML file using appropriate load options
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
        Workbook workbook = new Workbook(filePath, loadOptions);

        // Check whether the workbook contains a VBA project
        if (workbook.HasMacro)
        {
            // Access the VBA project and read its protection status
            bool isVbaProtected = workbook.VbaProject.IsProtected;

            // Output the result
            Console.WriteLine("VBA Project Protected: " + isVbaProtected);
        }
        else
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
        }
    }
}