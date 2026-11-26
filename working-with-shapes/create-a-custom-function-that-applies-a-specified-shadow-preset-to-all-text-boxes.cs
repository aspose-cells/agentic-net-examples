// Title: Apply a PresetShadowType to every TextBox shape in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: c# asp.net apply shadow preset to all text boxes in an Excel workbook using Aspose.Cells | set PresetShadowType for every TextBox shape across worksheets with Aspose.Cells .NET | batch change text box shadow effect in Excel files via Aspose.Cells C# example | enumerate TextBox collection and assign shadow effect in Aspose.Cells workbook | command line utility to update Excel text box shadows using Aspose.Cells library
// Tags: preset shadow assignment Aspose.Cells | text box shape shadow C# | iterate worksheet textboxes Aspose.Cells | enum parsing PresetShadowType C# | batch shape formatting Excel .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ApplyShadowToAllTextBoxes
{
    // The sample program validates command‑line arguments, loads an Excel workbook with Aspose.Cells, parses a PresetShadowType enum value (defaulting to OffsetDiagonalBottomRight), iterates through each worksheet and its TextBox collection, sets the ShadowEffect.PresetType for every TextBox, and saves the modified workbook. It includes error handling for missing files and invalid enum names.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Validate arguments
                if (args.Length < 2)
                {
                    Console.WriteLine("Usage: ApplyShadowToAllTextBoxes <input.xlsx> <output.xlsx> [ShadowPreset]");
                    Console.WriteLine("Available ShadowPreset values are defined in Aspose.Cells.Drawing.PresetShadowType enum.");
                    return;
                }

                string inputPath = args[0];
                string outputPath = args[1];
                string presetName = args.Length >= 3 ? args[2] : "OffsetDiagonalBottomRight";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                    return;
                }

                // Parse the preset enum value (case‑insensitive)
                if (!Enum.TryParse<PresetShadowType>(presetName, true, out PresetShadowType preset))
                {
                    Console.WriteLine($"Error: \"{presetName}\" is not a valid PresetShadowType.");
                    Console.WriteLine("Valid values are:");
                    foreach (var name in Enum.GetNames(typeof(PresetShadowType)))
                    {
                        Console.WriteLine($"  {name}");
                    }
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Loaded workbook \"{inputPath}\" successfully.");

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // TextBoxCollection implements IEnumerable, so we can use foreach
                    foreach (TextBox textBox in sheet.TextBoxes)
                    {
                        // Apply the shadow preset
                        textBox.ShadowEffect.PresetType = preset;
                        Console.WriteLine($"Applied shadow \"{preset}\" to TextBox on sheet \"{sheet.Name}\".");
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved with updated shadows to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
