using System;
using System.IO;
using Aspose.Cells;

namespace CombineWorkbooksDemo
{
    class Program
    {
        static void Main()
        {
            // Paths of the workbooks to be combined
            string[] sourceFiles = new string[]
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
            };

            // Create sample source workbooks if they do not exist
            for (int i = 0; i < sourceFiles.Length; i++)
            {
                if (!File.Exists(sourceFiles[i]))
                {
                    Workbook wb = new Workbook();
                    Worksheet ws = wb.Worksheets[0];
                    ws.Name = $"Sheet{i + 1}";
                    ws.Cells["A1"].PutValue($"Data from {sourceFiles[i]}");
                    wb.Save(sourceFiles[i], SaveFormat.Xlsx);
                }
            }

            // Create the destination workbook (empty workbook)
            Workbook combinedWorkbook = new Workbook();

            // Combine each source workbook into the destination workbook
            foreach (string filePath in sourceFiles)
            {
                Workbook sourceWorkbook = new Workbook(filePath);
                combinedWorkbook.Combine(sourceWorkbook);
            }

            // Save the combined workbook as XLSX
            combinedWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbooks have been successfully combined into 'CombinedWorkbook.xlsx'.");
        }
    }
}