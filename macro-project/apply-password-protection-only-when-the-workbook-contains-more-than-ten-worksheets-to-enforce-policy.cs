using System;
using Aspose.Cells;

namespace AsposeCellsPasswordPolicy
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add worksheets to demonstrate the condition (total 12 worksheets)
            for (int i = 0; i < 12; i++)
            {
                // The first worksheet already exists; add new ones for the rest
                if (i > 0)
                    workbook.Worksheets.Add();
                // Optionally put some data
                workbook.Worksheets[i].Cells["A1"].PutValue($"Sheet {i + 1}");
            }

            // Apply password protection only if the workbook has more than ten worksheets
            if (workbook.Worksheets.Count > 10)
            {
                // Protect the workbook structure with a password
                workbook.Protect(ProtectionType.Structure, "StrongPassword!123");
            }

            // Save the workbook
            workbook.Save("ProtectedIfMoreThanTenSheets.xlsx", SaveFormat.Xlsx);

            // Clean up
            workbook.Dispose();
        }
    }
}