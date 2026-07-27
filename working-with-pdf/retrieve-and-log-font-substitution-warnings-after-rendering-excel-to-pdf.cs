using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsFontWarningDemo
{
    // Author: Aspose.Cells .NET example – captures font substitution warnings during PDF rendering
    public class Program
    {
        public static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Prepare PDF save options with a custom warning callback
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Assign the custom callback to capture warnings
                WarningCallback = new RenderingWarningCallback()
            };

            // Save the workbook as PDF; warnings will be collected during this process
            workbook.Save("output.pdf", pdfOptions);

            // After saving, retrieve and log the total number of font substitution warnings captured
            RenderingWarningCallback callback = (RenderingWarningCallback)pdfOptions.WarningCallback;
            Console.WriteLine($"Total font substitution warnings captured: {callback.WarningCount}");
        }
    }

    // Custom implementation of IWarningCallback to handle rendering warnings
    public class RenderingWarningCallback : IWarningCallback
    {
        // Counter for font substitution warnings
        public int WarningCount { get; private set; }

        // This method is invoked by Aspose.Cells for each warning generated during rendering
        public void Warning(WarningInfo warningInfo)
        {
            // Filter only font substitution warnings
            if (warningInfo.WarningType == WarningType.FontSubstitution)
            {
                Console.WriteLine($"Font substitution warning: {warningInfo.Description}");
                WarningCount++;
            }
        }
    }
}