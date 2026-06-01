using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsReflectionDemo
{
    class Program
    {
        // Simple byte‑by‑byte comparison – returns true if files are identical
        static bool CompareByteArrays(byte[] a, byte[] b)
        {
            if (a == null || b == null) return false;
            if (a.Length != b.Length) return false;
            for (int i = 0; i < a.Length; i++)
                if (a[i] != b[i]) return false;
            return true;
        }

        static void Main()
        {
            try
            {
                // ---------- Create workbook and add a shape ----------
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];

                // Add a rectangle shape (position and size are arbitrary)
                Shape rect = ws.Shapes.AddRectangle(1, 1, 100, 100, 200, 100);

                // ---------- Render image BEFORE changing Reflection.Distance ----------
                string beforePath = "BeforeReflection.png";
                wb.Save(beforePath, SaveFormat.Png);

                // ---------- Set Reflection.Distance to 12 points ----------
                ReflectionEffect reflection = rect.Reflection;
                reflection.Type = ReflectionEffectType.HalfReflectionTouching; // make effect visible
                reflection.Distance = 12; // required setting
                reflection.Transparency = 0.5;
                reflection.Size = 80;
                reflection.Blur = 2;

                // ---------- Render image AFTER changing Reflection.Distance ----------
                string afterPath = "AfterReflection.png";
                wb.Save(afterPath, SaveFormat.Png);

                // ---------- Verify visual change via automated comparison ----------
                byte[] beforeBytes = File.Exists(beforePath) ? File.ReadAllBytes(beforePath) : Array.Empty<byte>();
                byte[] afterBytes = File.Exists(afterPath) ? File.ReadAllBytes(afterPath) : Array.Empty<byte>();
                bool imagesIdentical = CompareByteArrays(beforeBytes, afterBytes);
                Console.WriteLine("Images identical? " + imagesIdentical);
                Console.WriteLine(imagesIdentical
                    ? "Reflection distance change did NOT affect rendering."
                    : "Reflection distance change successfully altered rendering.");

                // ---------- Save workbook for manual inspection ----------
                string outputPath = "ReflectionDistanceDemo.xlsx";
                wb.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}