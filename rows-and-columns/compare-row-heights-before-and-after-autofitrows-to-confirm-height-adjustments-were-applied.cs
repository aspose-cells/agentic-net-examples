// Title: How to compare row heights before and after Worksheet.AutoFitRows in Aspose.Cells for .NET
// AI Prompts: Provide C# code that records each row's Height, invokes Worksheet.AutoFitRows, and logs the original and new heights to confirm adjustments using Aspose.Cells. | Generate a snippet that wraps long text, auto‑fits rows, captures row heights before and after AutoFitRows, and outputs a summary indicating which rows were resized.
// Common Searches: Aspose.Cells C# check if AutoFitRows changed row height | record original row height then auto fit rows Aspose.Cells example | compare worksheet row heights before and after AutoFitRows in .NET | detect row height adjustment after applying AutoFitRows with Aspose.Cells
// Tags: Aspose.Cells AutoFitRows row height verification | C# capture worksheet row heights Aspose.Cells | compare pre‑post AutoFitRows heights | log row height changes Aspose.Cells | auto‑fit rows with text wrapping Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, writes short and long wrapped text into cells A1‑A3, records each row’s initial Height, calls Worksheet.AutoFitRows(0,2), records the new heights, compares them, outputs whether each row was adjusted, and saves the file as AutoFitRowsResult.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells with text that will require row height adjustment
            sheet.Cells["A1"].PutValue("Short");
            sheet.Cells["A2"].PutValue("This is a very long text that should wrap and increase row height after AutoFitRows.");
            sheet.Cells["A3"].PutValue("Another long text that will need more height after autofit.");

            // Enable text wrapping for the cells that contain long text
            Style wrapStyle = workbook.CreateStyle();
            wrapStyle.IsTextWrapped = true;
            StyleFlag wrapFlag = new StyleFlag();
            wrapFlag.WrapText = true; // correct flag for text wrapping
            sheet.Cells["A2"].SetStyle(wrapStyle, wrapFlag);
            sheet.Cells["A3"].SetStyle(wrapStyle, wrapFlag);

            // Record original row heights (default is -1, meaning "auto")
            double[] originalHeights = new double[3];
            for (int i = 0; i < 3; i++)
            {
                originalHeights[i] = sheet.Cells.Rows[i].Height;
            }

            // Apply AutoFitRows to rows 0 through 2
            sheet.AutoFitRows(0, 2); // use overload without options

            // Record new row heights after AutoFitRows
            double[] newHeights = new double[3];
            for (int i = 0; i < 3; i++)
            {
                newHeights[i] = sheet.Cells.Rows[i].Height;
            }

            // Compare and output the results
            for (int i = 0; i < 3; i++)
            {
                bool adjusted = Math.Abs(originalHeights[i] - newHeights[i]) > 0.0001;
                Console.WriteLine($"Row {i + 1}: Original Height = {originalHeights[i]}, New Height = {newHeights[i]}, Adjusted = {adjusted}");
            }

            // Save the workbook (optional, just to verify the changes visually)
            string outputPath = "AutoFitRowsResult.xlsx";
            try
            {
                // Ensure the directory exists before saving
                string directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
