// Title: Determine VBA project protection with Aspose.Cells for .NET before editing an Excel workbook
// AI Prompts: Write a C# program that loads an .xlsx file using Aspose.Cells, checks workbook.VbaProject.IsProtected, and adds a new worksheet only when the VBA project is not protected. | Generate C# code that aborts any further workbook processing and logs a message if workbook.VbaProject.IsProtected returns true, otherwise continues with modifications.
// Common Searches: C# Aspose.Cells how to check if VBA project is password protected | skip adding worksheet when VBA macro is locked using Aspose.Cells .NET | use VbaProject.IsProtected property before modifying Excel with Aspose.Cells | detect VBA project protection status in Aspose.Cells C# example | prevent workbook save if VBA project is protected Aspose.Cells
// Tags: Aspose.Cells VBA protection detection | C# VbaProject.IsProtected usage | conditional Excel modification based on VBA lock | skip worksheet addition when VBA macro is protected | detect password-protected VBA project Aspose.Cells

using System;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, verifies the presence of a VBA project, reads its IsProtected flag, and proceeds to add a new worksheet and save the file only if the VBA project is not protected; otherwise it aborts processing and logs a message.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Check if the workbook contains a VBA project
        if (workbook.VbaProject != null)
        {
            // Use the IsProtected property to determine protection status
            bool isProtected = workbook.VbaProject.IsProtected;

            Console.WriteLine("VBA Project protection status: " + (isProtected ? "Protected" : "Not protected"));

            // Proceed only if the VBA project is not protected
            if (!isProtected)
            {
                // Place your further processing code here
                // Example: add a new worksheet
                int newSheetIndex = workbook.Worksheets.Add();
                workbook.Worksheets[newSheetIndex].Name = "NewSheet";

                // Save the modified workbook (replace with your desired output path)
                workbook.Save("output.xlsx");
                Console.WriteLine("Workbook processed and saved.");
            }
            else
            {
                Console.WriteLine("Processing aborted because the VBA project is protected.");
            }
        }
        else
        {
            Console.WriteLine("No VBA project found in the workbook.");
        }
    }
}
