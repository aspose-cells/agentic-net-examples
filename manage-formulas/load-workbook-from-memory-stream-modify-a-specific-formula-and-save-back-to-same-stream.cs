using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // ------------------------------------------------------------
            // 1. Create a workbook, add data and a formula, then save it
            //    to a MemoryStream. This simulates an existing Excel file
            //    that we later load from the same stream.
            // ------------------------------------------------------------
            using (MemoryStream stream = new MemoryStream())
            {
                Workbook wb = new Workbook();                         // create workbook
                Worksheet ws = wb.Worksheets[0];
                ws.Cells["A1"].PutValue(10);                         // sample data
                ws.Cells["A2"].PutValue(20);
                ws.Cells["B1"].Formula = "=SUM(A1:A2)";              // original formula
                wb.Save(stream, SaveFormat.Xlsx);                    // save to stream
                stream.Position = 0;                                 // reset for reading

                // ------------------------------------------------------------
                // 2. Load the workbook from the memory stream.
                // ------------------------------------------------------------
                Workbook loadedWb = new Workbook(stream);             // load from stream
                Worksheet loadedWs = loadedWb.Worksheets[0];

                // ------------------------------------------------------------
                // 3. Modify the specific formula (cell B1) to a new one.
                // ------------------------------------------------------------
                loadedWs.Cells["B1"].Formula = "=AVERAGE(A1:A2)";

                // Optional: recalculate formulas so the cell value reflects the change.
                loadedWb.CalculateFormula();

                // ------------------------------------------------------------
                // 4. Save the modified workbook back to the same MemoryStream.
                //    Clear the existing content first, then write the updated file.
                // ------------------------------------------------------------
                stream.SetLength(0);                                 // clear stream
                loadedWb.Save(stream, SaveFormat.Xlsx);              // save back to stream
                stream.Position = 0;                                 // reset for further use

                // ------------------------------------------------------------
                // 5. (Optional) Write the stream to a physical file to verify.
                // ------------------------------------------------------------
                string outputPath = "Modified.xlsx";
                using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    stream.CopyTo(file);
                }
            }
        }
        catch (Exception ex)
        {
            // Log or display the error as needed.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}