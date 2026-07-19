// Title: Dispose Aspose.Cells Workbook with C# using Statement for Immediate Resource Release
// Description: Shows how to create a Workbook inside a C# using block, add data, save as DisposedWorkbook.xlsx, and automatically call Dispose to free unmanaged resources.
// Keywords: Aspose.Cells workbook disposal | C# using statement | release unmanaged resources | Aspose.Cells memory management | Dispose Workbook Aspose | Excel file cleanup .NET | using block Aspose.Cells example
// Common Searches: how to dispose Aspose.Cells workbook in C# | C# using statement for Aspose.Cells | release unmanaged resources after saving Excel with Aspose | prevent memory leaks Aspose.Cells | automatic workbook cleanup .NET
// Developer Intent: Promptly free native resources by disposing the Aspose.Cells Workbook through a using block.
// Use Cases: Create a temporary Excel file, save it, and rely on the using block for cleanup. | Wrap multiple workbook operations in a using statement to guarantee disposal even on exceptions. | Implement resource‑efficient Excel generation in server‑side .NET services.
// AI Prompts: Write C# code that creates an Aspose.Cells Workbook, writes data, saves it, and ensures disposal with a using block. | Explain why disposing an Aspose.Cells Workbook promptly is critical for unmanaged resource management. | List best practices for managing the lifecycle of Aspose.Cells objects in .NET applications.

using Aspose.Cells;
using System;

namespace Example
{
    // Shows how to create a Workbook inside a C# using block, add data, save as DisposedWorkbook.xlsx, and automatically call Dispose to free unmanaged resources.
    class Program
    {
        static void Main()
        {
            // Workbook is created inside a using block.
            // The using statement ensures Dispose() is called automatically.
            using (Workbook workbook = new Workbook())
            {
                // Access the default worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Add some sample data.
                sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");

                // Save the workbook to a file.
                workbook.Save("DisposedWorkbook.xlsx");
            } // workbook.Dispose() is invoked here.
        }
    }
}
