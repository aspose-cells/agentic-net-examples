using System;
using Aspose.Cells;

namespace AsposeCellsProtectionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the protection object of the worksheet
            Protection protection = worksheet.Protection;

            // Allow sorting on the protected sheet
            protection.AllowSorting = true;

            // Disallow filtering on the protected sheet
            protection.AllowFiltering = false;

            // Apply protection to the worksheet (protect all aspects)
            worksheet.Protect(ProtectionType.All);

            // Save the workbook to verify the protection settings
            workbook.Save("WorksheetProtected_SortAllowed_FilterBlocked.xlsx");
        }
    }
}