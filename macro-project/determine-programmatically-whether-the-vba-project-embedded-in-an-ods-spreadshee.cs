using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Path to the ODS spreadsheet
        string odsPath = "sample.ods";

        // Load the ODS workbook
        Workbook workbook = new Workbook(odsPath);

        // Verify that the workbook actually contains a VBA project
        if (workbook.HasMacro)
        {
            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Determine whether the VBA project is protected
            bool isProtected = vbaProject.IsProtected;

            Console.WriteLine("VBA Project Protected: " + isProtected);
        }
        else
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
        }
    }
}