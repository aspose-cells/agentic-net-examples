// Title: Aspose.Cells .NET – Compare HTML file size with WidthScalable true vs false
// Description: Creates a workbook, fills it with sample data, saves it to HTML twice (WidthScalable = true and WidthScalable = false), measures each file’s byte size, and reports which setting generates a larger HTML output.
// Keywords: Aspose.Cells | HtmlSaveOptions | WidthScalable | HTML export size | C# .NET | file size comparison | Excel to HTML | workbook rendering
// Common Searches: WidthScalable impact on HTML size Aspose.Cells | Aspose.Cells HTML file size true vs false | measure HTML output size C# Aspose | compare WidthScalable settings Aspose.Cells | optimize HTML export size Aspose.Cells
// Developer Intent: Find out whether enabling or disabling HtmlSaveOptions.WidthScalable produces a larger HTML file for the same workbook.
// Use Cases: Assess the bandwidth cost of scalable HTML layouts before publishing Excel data on a website. | Select the most size‑efficient WidthScalable setting for automated report generation pipelines. | Benchmark HTML export options across multiple workbooks to define a default export configuration.
// AI Prompts: Write C# code that saves a workbook to HTML with WidthScalable true and false, then prints the byte sizes and indicates which is larger. | Explain why the WidthScalable option can increase HTML file size and suggest techniques to keep the output compact while preserving scalability. | Create a PowerShell script that processes a folder of .xlsx files, exports each to HTML with both WidthScalable values, and logs the size differences to a CSV file.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, fills it with sample data, saves it to HTML twice (WidthScalable = true and WidthScalable = false), measures each file’s byte size, and reports which setting generates a larger HTML output.
class CompareWidthScalable
{
    static void Main()
    {
        // Create a workbook and populate it with sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add enough data to make the HTML size noticeable
        for (int row = 0; row < 200; row++)
        {
            for (int col = 0; col < 20; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row}C{col}");
            }
        }

        // Initialize HTML save options (common for both saves)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Save with WidthScalable = true
        htmlOptions.WidthScalable = true;
        string truePath = "output_widthscalable_true.html";
        workbook.Save(truePath, htmlOptions);
        long trueSize = new FileInfo(truePath).Length;

        // Save with WidthScalable = false
        htmlOptions.WidthScalable = false;
        string falsePath = "output_widthscalable_false.html";
        workbook.Save(falsePath, htmlOptions);
        long falseSize = new FileInfo(falsePath).Length;

        // Output the file sizes and comparison result
        Console.WriteLine($"WidthScalable = true  : {trueSize} bytes");
        Console.WriteLine($"WidthScalable = false : {falseSize} bytes");

        if (trueSize > falseSize)
            Console.WriteLine("Enabling WidthScalable results in a larger HTML file.");
        else if (trueSize < falseSize)
            Console.WriteLine("Disabling WidthScalable results in a larger HTML file.");
        else
            Console.WriteLine("Both HTML files have the same size.");
    }
}
