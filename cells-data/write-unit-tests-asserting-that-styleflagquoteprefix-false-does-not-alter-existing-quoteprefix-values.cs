using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Access cell A1 and put a numeric string (no leading quote)
                Cell cell = sheet.Cells["A1"];
                cell.PutValue("12345");

                // Create a style that enables QuotePrefix (treats the value as text)
                Style styleWithPrefix = workbook.CreateStyle();
                styleWithPrefix.QuotePrefix = true;

                // Create a StyleFlag that enables the QuotePrefix flag
                StyleFlag flagWithPrefix = new StyleFlag();
                flagWithPrefix.QuotePrefix = true;

                // Apply the style to the cell – this should set QuotePrefix to true
                cell.SetStyle(styleWithPrefix, flagWithPrefix);
                if (!cell.GetStyle().QuotePrefix)
                    throw new Exception("QuotePrefix should be true after first style application.");

                // Now create a new style where QuotePrefix is false
                Style styleWithoutPrefix = workbook.CreateStyle();
                styleWithoutPrefix.QuotePrefix = false;

                // Create a StyleFlag where QuotePrefix flag is false (property ignored)
                StyleFlag flagWithoutPrefix = new StyleFlag();
                flagWithoutPrefix.QuotePrefix = false;

                // Apply the new style – existing QuotePrefix should remain unchanged
                cell.SetStyle(styleWithoutPrefix, flagWithoutPrefix);
                if (!cell.GetStyle().QuotePrefix)
                    throw new Exception("QuotePrefix should remain true when StyleFlag.QuotePrefix is false.");

                // Save to a memory stream (no file I/O)
                using (var ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsx);
                    if (ms.Length == 0)
                        throw new Exception("Workbook should be saved to the memory stream.");
                }

                Console.WriteLine("Test completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}