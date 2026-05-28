using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace PivotRefreshAndPdfExport
{
    class Program
    {
        static void Main()
        {
            // List of source Excel files to process
            string[] sourceFiles = new string[]
            {
                "Report1.xlsx",
                "Report2.xlsx",
                "Report3.xlsx"
            };

            foreach (string sourcePath in sourceFiles)
            {
                // Load the workbook (creation & loading)
                Workbook workbook = new Workbook(sourcePath);

                // Refresh all pivot tables in all worksheets
                workbook.Worksheets.RefreshPivotTables();

                // Save the refreshed workbook to a temporary XLSX file
                string refreshedPath = Path.Combine(
                    Path.GetDirectoryName(sourcePath) ?? string.Empty,
                    Path.GetFileNameWithoutExtension(sourcePath) + "_refreshed.xlsx");

                workbook.Save(refreshedPath, SaveFormat.Xlsx);

                // Define the output PDF file name
                string pdfPath = Path.ChangeExtension(sourcePath, ".pdf");

                // Convert the refreshed XLSX to PDF (saving)
                ConversionUtility.Convert(refreshedPath, pdfPath);

                // Clean up the temporary refreshed file
                if (File.Exists(refreshedPath))
                {
                    File.Delete(refreshedPath);
                }

                Console.WriteLine($"Processed '{sourcePath}' -> '{pdfPath}'");
            }
        }
    }
}