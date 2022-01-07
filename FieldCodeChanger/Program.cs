using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using FieldCodeChanger.DTO;
using FieldCodeChanger.ExtensionMethods;

namespace FieldCodeChanger
{
    internal static class Program
    {
        private static void Main()
        {
            var changes = new List<FieldCodeManipulationDto>
            {
                new(@"{ FIELDCODE1 }/{ FIELDCODE2 }/{ FIELDCODE3 }", @"{ FIELDCODE2 }/{ FIELDCODE3 }/{ FIELDCODE1 }")
            };

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            const string fileToOpen = @"C:\\Users\georg\Documents\WordDoc\test.docx";
            using var doc = WordprocessingDocument.Open(fileToOpen, false);
            OpenXmlPart part = doc.MainDocumentPart;
            Field.Field.AnnotateWithFieldInfo(part);
            var root = part.GetXDocument().Root;
            if (root != null)
            {
                var maxFieldId =
                    root.Descendants().Select(e =>
                    {
                        var stack = e.Annotation<Stack<Field.Field.FieldElementTypeInfo>>();
                        return stack != null ? stack.Select(s => s.Id).Max() : 0;
                    }).Max();

                for (var id = 1; id <= maxFieldId; ++id) Console.WriteLine($"{id}: {Field.Field.InstrText(root, id)}");
            }

            Console.WriteLine("=======================================");
            var xElement = part.GetXDocument().Root;
            if (xElement == null) return;
            {
                foreach (var item in xElement.Descendants())
                {
                    Console.Write(
                        $"{item.Name.LocalName,-20}{(item.Name == w + "fldChar" ? item.Attribute(w + "fldCharType")?.Value.PadRight(16) : "".PadRight(16))}");

                    var stack = item.Annotation<Stack<Field.Field.FieldElementTypeInfo>>();
                    if (stack == null) continue;

                    foreach (var item2 in stack)
                    {
                        Console.Write($"{item2.Id.ToString(),-4}:{item2.FieldElementType.ToString(),-16}");
                        Console.WriteLine();
                    }
                }
            }
        }
    }
}