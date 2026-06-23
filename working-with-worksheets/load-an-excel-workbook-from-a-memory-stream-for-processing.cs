using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class LoadWorkbookFromMemoryStream
    {
        public static void Run()
        {
            try
            {
                // 1. Create a sample workbook and add some data.
                using (Workbook sourceWorkbook = new Workbook())
                {
                    Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
                    sourceSheet.Cells["A1"].PutValue("Item");
                    sourceSheet.Cells["B1"].PutValue("Quantity");
                    sourceSheet.Cells["A2"].PutValue("Apple");
                    sourceSheet.Cells["B2"].PutValue(10);
                    sourceSheet.Cells["A3"].PutValue("Banana");
                    sourceSheet.Cells["B3"].PutValue(20);

                    // 2. Save the workbook to a MemoryStream.
                    using (MemoryStream memoryStream = sourceWorkbook.SaveToStream())
                    {
                        memoryStream.Position = 0; // Reset position before reading.

                        // 3. Load a new workbook from the MemoryStream.
                        using (Workbook loadedWorkbook = new Workbook(memoryStream))
                        {
                            // 4. Process the loaded workbook (example: read a cell value).
                            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                            string item = loadedSheet.Cells["A2"].StringValue;
                            double quantity = loadedSheet.Cells["B2"].DoubleValue;

                            Console.WriteLine($"Loaded data - Item: {item}, Quantity: {quantity}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application.
    public class Program
    {
        public static void Main(string[] args)
        {
            LoadWorkbookFromMemoryStream.Run();
        }
    }
}