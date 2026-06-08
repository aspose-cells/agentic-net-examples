using System;
using Aspose.Cells;

namespace AsposeCellsExportXlsb
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["B1"].PutValue(123);

            // Protect the workbook structure with a password
            // This ensures that the workbook's structure (e.g., adding/removing sheets) remains protected
            workbook.Protect(ProtectionType.Structure, "myPassword");

            // Create XLSB save options
            XlsbSaveOptions saveOptions = new XlsbSaveOptions
            {
                // Export all column indexes (default is true, set explicitly for clarity)
                ExportAllColumnIndexes = true
            };

            // Save the workbook as an XLSB file while retaining the protection
            workbook.Save("ProtectedWorkbook.xlsb", saveOptions);
        }
    }
}