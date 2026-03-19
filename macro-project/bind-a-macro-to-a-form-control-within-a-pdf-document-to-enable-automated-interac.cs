using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfMacroDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a button (form control) to the worksheet
            // Parameters: upper left row, upper left column, upper left pixel offset Y, upper left pixel offset X, lower right row, lower right column
            Button button = (Button)sheet.Shapes.AddButton(2, 2, 0, 0, 4, 4);
            button.Text = "Run Macro";

            // Bind a macro to the button using the MacroName property
            button.MacroName = "MyMacro()";

            // (Optional) Define the macro in the workbook's VBA project
            // Note: In a real scenario, you would add VBA code to the workbook.
            // For demonstration, we only set the macro name.

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF; the button will retain the macro binding
            workbook.Save("FormWithMacro.pdf", pdfOptions);

            Console.WriteLine("PDF with macro-bound button created successfully.");
        }
    }
}