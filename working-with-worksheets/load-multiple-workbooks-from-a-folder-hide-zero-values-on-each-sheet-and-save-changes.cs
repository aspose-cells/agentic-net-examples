using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchProcessing
{
    class Program
    {
        static void Main()
        {
            // Folder containing the Excel files
            string folderPath = @"C:\ExcelFiles";

            // Get all Excel files in the folder (xlsx, xls, xlsm, etc.)
            string[] excelFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in excelFiles)
            {
                // Load the workbook using the constructor that accepts a file path
                Workbook workbook = new Workbook(filePath);

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Hide zero values by disabling the DisplayZeros property
                    sheet.DisplayZeros = false;
                }

                // Save the workbook back to the same file (overwrites the original)
                workbook.Save(filePath);
            }

            Console.WriteLine("Processing completed for all workbooks in the folder.");
        }
    }
}