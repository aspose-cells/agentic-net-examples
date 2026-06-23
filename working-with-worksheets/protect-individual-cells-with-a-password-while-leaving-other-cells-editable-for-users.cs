using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ProtectIndividualCellWithPassword
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Unlock all cells so they are editable when the sheet is protected
            Style unlockedStyle = workbook.CreateStyle();
            unlockedStyle.IsLocked = false;
            StyleFlag flag = new StyleFlag
            {
                Locked = true // apply IsLocked property
            };
            cells.ApplyStyle(unlockedStyle, flag);

            // Define a range that should be protected with a password (cell B2)
            int rangeIndex = sheet.AllowEditRanges.Add("ProtectedCellB2", 1, 1, 1, 1);
            ProtectedRange protectedRange = sheet.AllowEditRanges[rangeIndex];
            protectedRange.Password = "cellpwd";

            // Protect the worksheet (no sheet password needed)
            sheet.Protect(ProtectionType.All);

            // Determine output file path
            string outputPath = "ProtectedIndividualCell.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
    }
}