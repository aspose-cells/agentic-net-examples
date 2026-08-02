// Title: Unprotect a worksheet and unfreeze frozen panes using Aspose.Cells for .NET (C#)
// Description: Shows how to protect a worksheet with a password, freeze panes at a given cell, then unprotect the sheet and call UnFreezePanes to release the frozen rows and columns before saving the workbook. The example illustrates the required call order for Aspose.Cells in C#.
// Keywords: Aspose.Cells C# unprotect worksheet | Aspose.Cells UnFreezePanes | freeze panes protected worksheet .NET | Worksheet.Protect Aspose.Cells | remove frozen rows columns Aspose.Cells | Aspose.Cells workbook protection | C# Excel unfreeze panes after protection
// Common Searches: C# Aspose.Cells unfreeze panes on protected sheet | How to unprotect worksheet before UnFreezePanes Aspose.Cells | Aspose.Cells unfreeze frozen rows after protection | Unfreeze panes Aspose.Cells .NET example | Remove frozen panes from protected worksheet using Aspose.Cells
// Developer Intent: Programmatically unprotect a protected worksheet so that UnFreezePanes can be executed successfully.
// Use Cases: Temporarily lift protection to modify or clear frozen panes during automated report generation. | Reset view settings in a template that was previously protected and had frozen rows/columns. | Prepare a workbook for distribution by removing protection and clearing frozen panes in one step.
// AI Prompts: Generate C# code with Aspose.Cells that protects a worksheet, freezes panes at a specific cell, then safely unprotects and unfreezes them, including error handling. | Explain why Aspose.Cells requires a worksheet to be unprotected before calling UnFreezePanes and provide the correct sequence of API calls. | Create a reusable C# method that accepts a workbook path, password, and cell address, ensures the sheet is unprotected, and then unfreezes any frozen panes.

using System;
using Aspose.Cells;

namespace AsposeCellsUnprotectAndUnfreeze
{
    // Shows how to protect a worksheet with a password, freeze panes at a given cell, then unprotect the sheet and call UnFreezePanes to release the frozen rows and columns before saving the workbook. The example illustrates the required call order for Aspose.Cells in C#.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Protect the worksheet with a password
            worksheet.Protect(ProtectionType.All, "myPassword", null);
            Console.WriteLine("Worksheet protected: " + worksheet.IsProtected);

            // Freeze panes at cell C3 (rows 2, columns 2 frozen)
            worksheet.FreezePanes("C3", 2, 2);
            Console.WriteLine("Panes frozen.");

            // Unprotect the worksheet (required before unfreezing)
            worksheet.Unprotect("myPassword");
            Console.WriteLine("Worksheet unprotected: " + !worksheet.IsProtected);

            // Unfreeze the panes
            worksheet.UnFreezePanes();
            Console.WriteLine("Panes unfrozen.");

            // Save the workbook
            workbook.Save("UnprotectAndUnfreezeDemo.xlsx");
        }
    }
}
