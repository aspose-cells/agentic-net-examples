using System;
using Aspose.Cells;

namespace AsposeCellsWriteProtectionDemo
{
    public class Program
    {
        public static void Main()
        {
            // Path to the existing workbook
            string inputPath = "input.xlsx";

            // Path where the modified workbook will be saved
            string outputPath = "output.xlsx";

            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(inputPath);

            // Get the current Windows user name
            string currentUser = Environment.UserName;

            // Change the write‑protection author to the current user
            workbook.Settings.WriteProtection.Author = currentUser;

            // Re‑save the workbook (uses the provided save rule)
            workbook.Save(outputPath);
        }
    }
}