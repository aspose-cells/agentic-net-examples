// Title: C# Example: Unprotect a Worksheet and Unfreeze Panes with Aspose.Cells
// Description: Demonstrates how to create a workbook, protect the first worksheet with a password, freeze panes at cell C3, then safely unprotect the sheet and call UnFreezePanes before saving the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# unprotect worksheet | UnFreezePanes Aspose.Cells | freeze panes C# Aspose.Cells | worksheet protection API Aspose | remove worksheet protection before unfreeze | Aspose.Cells workbook example | C# Excel unprotect and unfreeze | Aspose.Cells sample code | protect and unfreeze panes | Aspose.Cells API usage
// Common Searches: C# Aspose.Cells unprotect worksheet before UnFreezePanes | how to unfreeze frozen panes after removing protection Aspose.Cells | Aspose.Cells example protect freeze unprotect unfreeze | unprotect worksheet Aspose.Cells C# password | unfreeze panes Aspose.Cells after unprotect
// Developer Intent: The developer needs to remove worksheet protection prior to calling UnFreezePanes so the unfreeze operation succeeds.
// Use Cases: Automated workflow that unlocks a protected template and removes frozen panes before further processing. | Preparing a workbook for printing or export by clearing protection and pane freezes programmatically. | Dynamic report generation where a sheet is initially secured and later opened for editing without layout constraints.
// AI Prompts: Generate C# code using Aspose.Cells that checks if a worksheet is protected, unprotects it with a given password, and then safely calls UnFreezePanes. | Provide an Aspose.Cells snippet that protects a worksheet, freezes panes at a specific cell, and later unprotects and unfreezes the sheet while handling possible exceptions.

using System;
using Aspose.Cells;

namespace AsposeCellsUnprotectAndUnfreeze
{
    // Demonstrates how to create a workbook, protect the first worksheet with a password, freeze panes at cell C3, then safely unprotect the sheet and call UnFreezePanes before saving the file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Protect the worksheet with a password
            worksheet.Protect(ProtectionType.All, "myPassword", null);
            Console.WriteLine($"Worksheet protected: {worksheet.IsProtected}");

            // Freeze panes at cell C3 (rows 2, columns 2 frozen) to set up a scenario
            worksheet.FreezePanes("C3", 2, 2);
            Console.WriteLine("Panes frozen.");

            // Unprotect the worksheet using the correct password
            worksheet.Unprotect("myPassword");
            Console.WriteLine($"Worksheet protected after unprotect: {worksheet.IsProtected}");

            // Unfreeze the panes now that the sheet is unprotected
            worksheet.UnFreezePanes();
            Console.WriteLine("Panes unfrozen.");

            // Save the workbook (lifecycle: save)
            workbook.Save("UnprotectAndUnfreezeDemo.xlsx");
            Console.WriteLine("Workbook saved.");
        }
    }
}
