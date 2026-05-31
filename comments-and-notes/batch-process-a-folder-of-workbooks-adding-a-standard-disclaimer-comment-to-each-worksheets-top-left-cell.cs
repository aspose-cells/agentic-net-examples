using System;
using System.IO;
using Aspose.Cells;

namespace BatchDisclaimer
{
    public static class WorkbookProcessor
    {
        // Adds a disclaimer comment to the top‑left cell (A1) of every worksheet in all workbooks
        // found in the specified folder. The original files are overwritten with the updated version.
        public static void AddDisclaimerToFolder(string folderPath, string disclaimer)
        {
            // Get all Excel files in the folder (you can adjust the pattern if needed)
            string[] files = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string filePath in files)
            {
                // Load the workbook using the constructor that accepts a file path
                using (Workbook workbook = new Workbook(filePath))
                {
                    // Iterate through each worksheet in the workbook
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Add a comment to cell A1 (top‑left cell)
                        int commentIndex = sheet.Comments.Add("A1");
                        sheet.Comments[commentIndex].Note = disclaimer;
                    }

                    // Save the modified workbook back to the same file (overwrite)
                    workbook.Save(filePath);
                }
            }
        }

        // Example usage
        public static void Main()
        {
            string folder = @"C:\ExcelFiles"; // Folder containing the workbooks
            string disclaimerText = "This workbook is confidential and intended for authorized personnel only.";

            AddDisclaimerToFolder(folder, disclaimerText);

            Console.WriteLine("Disclaimer added to all workbooks in the folder.");
        }
    }
}