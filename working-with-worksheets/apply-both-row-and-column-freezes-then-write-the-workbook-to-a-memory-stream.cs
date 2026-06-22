using System;
using System.IO;
using Aspose.Cells;

namespace Demo
{
    public class FreezeAndSaveDemo
    {
        // Returns a MemoryStream containing the workbook with frozen rows and columns
        public static MemoryStream Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data
                sheet.Cells["A1"].PutValue("Header1");
                sheet.Cells["B1"].PutValue("Header2");
                sheet.Cells["A2"].PutValue("Data1");
                sheet.Cells["B2"].PutValue("Data2");

                // Freeze first 2 rows and first 2 columns (freeze point at C3)
                sheet.FreezePanes(2, 2, 2, 2);

                // Save to a memory stream (Excel 97‑2003 format)
                MemoryStream stream = workbook.SaveToStream();
                stream.Position = 0;
                return stream;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error in FreezeAndSaveDemo: {ex.Message}");
                throw;
            }
        }
    }

    public class Program
    {
        // Entry point required for compilation
        public static void Main(string[] args)
        {
            try
            {
                using (MemoryStream ms = FreezeAndSaveDemo.Run())
                {
                    string outputPath = "FrozenWorkbook.xls";

                    // Ensure the directory exists
                    string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                    if (!Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    // Write the stream to a file
                    using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        ms.CopyTo(file);
                    }

                    Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}