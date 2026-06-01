using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace SmartArtReport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the Excel files to analyze
            string folderPath = @"C:\ExcelFiles";

            // Get all .xlsx files in the folder
            string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx");

            Console.WriteLine("Workbook,Worksheet,SmartArtShapeCount");

            foreach (string filePath in excelFiles)
            {
                // Load the workbook (using the standard load constructor)
                Workbook workbook = new Workbook(filePath);

                // Extract the workbook file name
                string workbookName = Path.GetFileName(filePath);

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Count shapes that are SmartArt
                    int smartArtCount = sheet.Shapes
                                            .Cast<Shape>()
                                            .Count(shape => shape.IsSmartArt);

                    // Output the result as CSV line
                    Console.WriteLine($"{workbookName},{sheet.Name},{smartArtCount}");
                }
            }
        }
    }
}