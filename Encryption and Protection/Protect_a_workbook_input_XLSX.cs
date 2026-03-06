using System;
using Aspose.Cells;

namespace WorkbookProtectionExample
{
    class Program
    {
        static void Main()
        {
            // Load the existing workbook (input XLSX)
            Workbook workbook = new Workbook("input.xlsx");

            // Protect the workbook's structure with a password
            workbook.Protect(ProtectionType.Structure, "mySecretPassword");

            // Save the protected workbook
            workbook.Save("protected_output.xlsx", SaveFormat.Xlsx);

            // Release resources
            workbook.Dispose();
        }
    }
}