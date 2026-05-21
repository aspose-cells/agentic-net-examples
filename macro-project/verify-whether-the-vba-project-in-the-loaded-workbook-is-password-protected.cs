using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCheck
{
    class Program
    {
        static void Main()
        {
            // Load the workbook that may contain a VBA project
            Workbook workbook = new Workbook("input.xlsm");

            // Get the VBA project from the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Determine if the VBA project is protected with a password
            bool isPasswordProtected = vbaProject.IsProtected;

            // Output the verification result
            Console.WriteLine("VBA Project password protected: " + isPasswordProtected);
        }
    }
}