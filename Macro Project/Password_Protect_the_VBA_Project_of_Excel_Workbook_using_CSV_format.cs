using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtection
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source CSV file
            string csvPath = "input.csv";

            // Load the CSV file into a workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // Protect the VBA project (lock for viewing) with a password
            // The first argument 'true' locks the project for viewing
            workbook.VbaProject.Protect(true, "MyVbaPassword");

            // Save the workbook as a macro‑enabled Excel file (XLSM)
            workbook.Save("ProtectedVbaProject.xlsm", SaveFormat.Xlsm);
        }
    }
}