// Title: Load an Excel workbook from a Stream, assign a VBA macro to a shape control, and save the result to a MemoryStream with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells code to attach a VBA macro to a button shape | reading an Excel file from a Stream and updating shape hyperlink with Aspose.Cells | saving an edited Excel workbook to a MemoryStream using Aspose.Cells | how to set shape.Hyperlink.Address to a macro name in Aspose.Cells | find a shape by its name on a worksheet and assign a macro in C#
// Tags: assign macro to shape Aspose.Cells | load workbook from stream C# Aspose.Cells | save workbook to memory stream Aspose.Cells | shape hyperlink address macro Aspose.Cells | retrieve shape by name worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example demonstrates how to load an Excel workbook from a .NET Stream, locate a shape (such as a button) by its name on the first worksheet, assign a VBA macro to the shape by setting its Hyperlink.Address, and then save the modified workbook into a MemoryStream using Aspose.Cells for .NET.
public class MacroAssigner
{
    /// <param name="inputStream">Stream containing the source workbook.</param>
    /// <param name="controlName">Name of the control (shape) to which the macro will be assigned.</param>
    /// <param name="macroName">Name of the macro to assign to the control.</param>
    /// <returns>A MemoryStream containing the saved workbook.</returns>
    public MemoryStream AssignMacroToControl(Stream inputStream, string controlName, string macroName)
    {
        try
        {
            // Load the workbook from the input stream
            var workbook = new Workbook(inputStream);

            // Locate the control (shape) by name on the first worksheet
            var worksheet = workbook.Worksheets[0];
            Shape shape = worksheet.Shapes[controlName];

            if (shape == null)
                throw new ArgumentException($"Control with name '{controlName}' was not found on the worksheet.");

            // Assign the macro to the control via its Hyperlink object
            // The Hyperlink property is read‑only, but its members can be modified.
            shape.Hyperlink.Address = macroName; // Macro name is stored in the address field

            // Save the modified workbook into a memory stream
            var outputStream = new MemoryStream();
            workbook.Save(outputStream, SaveFormat.Xlsx);
            outputStream.Position = 0; // Reset for reading

            return outputStream;
        }
        catch (Exception)
        {
            // Rethrow to let the caller handle the exception
            throw;
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string controlName = "MyButton";
            const string macroName = "MyMacro";

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            using (FileStream inputStream = File.OpenRead(inputPath))
            {
                var assigner = new MacroAssigner();
                using (MemoryStream resultStream = assigner.AssignMacroToControl(inputStream, controlName, macroName))
                {
                    File.WriteAllBytes(outputPath, resultStream.ToArray());
                    Console.WriteLine($"Workbook saved to '{outputPath}'.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
