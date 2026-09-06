// Title: C# example that references the Excel COM library, writes and reads a cell, releases COM objects, and then creates an Aspose.Cells workbook
// AI Prompts: Create a static C# class named ExcelComHelper that instantiates the Excel.Application COM object via Type.GetTypeFromProgID, writes a supplied string to cell A1, reads the value back, closes the workbook, quits Excel, and releases the COM object with Marshal.ReleaseComObject. | Add error handling that verifies the Excel COM ProgID is registered and throws an InvalidOperationException with a clear message if it is missing. | Write a C# console program that calls ExcelComHelper.WriteAndReadCell, prints the returned value, then uses Aspose.Cells to create a new workbook, puts text into cell B2, and saves the file as Output.xlsx. | Show how to add a COM reference to Excel in a .csproj file so the code compiles without requiring the primary interop assembly.
// Common Searches: how to automate Excel from C# using COM without installing the primary interop assembly | C# write to cell A1 and read it back using dynamic Excel.Application COM object | proper way to release Excel COM objects in .NET to prevent memory leaks | combine Aspose.Cells workbook generation with Excel COM interop in a single C# project | add COM reference to Excel in .NET Core project file for compile-time support
// Tags: excel com automation c# | dynamic excel.application interop | marshal.releasecomobject cleanup excel | aspose.cells workbook creation c# | excel com write read cell example | csproj com reference configuration

using System;
using System.Runtime.InteropServices;

namespace AsposeCellsComInteropDemo
{
    // This module demonstrates referencing an external COM library (Microsoft Excel) 
    // and using it from .NET code. Ensure that the COM library is registered on the machine.
    // Demonstrates referencing the Excel COM library in C#, writing and reading a value in cell A1 via dynamic COM, correctly releasing COM resources with Marshal.ReleaseComObject, and then creating an Aspose.Cells workbook, inserting text into B2, and saving it as Output.xlsx.
    public static class ExcelComHelper
    {
        // Creates an instance of the Excel Application COM object, writes a value to a cell,
        // reads it back, and then quits the application.
        public static string WriteAndReadCell(string valueToWrite)
        {
            // ProgID for Excel Application
            const string progId = "Excel.Application";

            // Create COM object
            Type excelType = Type.GetTypeFromProgID(progId);
            if (excelType == null)
                throw new InvalidOperationException("Excel COM component is not registered.");

            // Use dynamic to avoid needing the primary interop assembly at compile time
            dynamic excelApp = null;
            try
            {
                excelApp = Activator.CreateInstance(excelType);
                excelApp.Visible = false; // Keep Excel hidden

                // Add a new workbook and get the first worksheet
                dynamic workbook = excelApp.Workbooks.Add();
                dynamic worksheet = workbook.Worksheets[1];

                // Write the value to cell A1
                worksheet.Range["A1"].Value = valueToWrite;

                // Read the value back from cell A1
                string readValue = worksheet.Range["A1"].Value?.ToString() ?? string.Empty;

                // Clean up
                workbook.Close(false);
                excelApp.Quit();

                return readValue;
            }
            finally
            {
                // Release COM objects to prevent memory leaks
                if (excelApp != null)
                {
                    Marshal.ReleaseComObject(excelApp);
                }
            }
        }
    }

    // Example usage with Aspose.Cells (optional demonstration)
    class Program
    {
        static void Main()
        {
            // Use the COM helper to write and read a value
            string testValue = "Hello from COM!";
            string result = ExcelComHelper.WriteAndReadCell(testValue);
            Console.WriteLine($"COM returned: {result}");

            // Aspose.Cells example to create a workbook (uses provided lifecycle rules)
            var workbook = new Aspose.Cells.Workbook(); // create
            var sheet = workbook.Worksheets[0];
            sheet.Cells["B2"].PutValue("Aspose.Cells works!");
            workbook.Save("Output.xlsx"); // save
        }
    }
}
