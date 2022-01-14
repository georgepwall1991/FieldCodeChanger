using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using FieldCodeChanger.DTO;
using FieldCodeChanger.ExtensionMethods;

namespace FieldCodeChanger //git test #47, the robots are starting to suspect I am not one of them.
{
    internal static class Program
    {
        private static void Main()
        {
            var changes = new List<FieldCodeManipulationDto>
            {
                new(@"{FEINIT}{CLIENTNO}{MATTERNO}", @"{CLIENTNO}/{MATTERNO}/{FEINIT \*charformat}")
                new(@"{FEINIT \*charformat}/{CLIENTNO \*arabic}/{MATTERNO \*arabic}", @"{CLIENTNO \*arabic}/{MATTERNO \*arabic}/{FEINIT \*charformat}"
                
                /*
                 * What target sequences should all be changed to:
                 * { CLIENTNO \* arabic \*charformat }/{ MATTERNO \* arabic \*charformat }/{ FEINIT \*charformat }
                 *
                 * List of target sequences (all possible cases to be included):
                 * {FEINIT}{CLIENTNO}{MATTERNO}
                 * { FEINIT }{ CLIENTNO }{ MATTERNO }
                 * {FEINIT}/{CLIENTNO}/{MATTERNO}
                 * { FEINIT }/{ CLIENTNO }/{ MATTERNO }
                 * {FEINIT}/{CLIENTNO \*arabic}/{MATTERNO \*arabic}
                 * { FEINIT }/{ CLIENTNO \*arabic }/{ MATTERNO \*arabic }
                 * {CLIENTNO}/{MATTERNO}/{FEINIT \*charformat}
                 * {FEINIT \*charformat}/{CLIENTNO \*arabic}/{MATTERNO \*arabic}
                 * {CLIENTNO \*arabic}/{MATTERNO \*arabic}/{FEINIT \*charformat}
                 * {FEINIT \*charformat}/{CLIENTNO \* arabic}/{MATTERNO \* arabic}
                 * { FEINIT \*charformat }/{ CLIENTNO \* arabic }/{ MATTERNO \* arabic }
                 * { FEINIT \*charformat }/{ CLIENTNO \* arabic \*charformat }/{ MATTERNO \* arabic \*charformat }
                 * {FEINIT \*charformat}{CLIENTNO \* arabic}{MATTERNO \* arabic}
                 * { FEINIT \*charformat }{ CLIENTNO \* arabic }{ MATTERNO \* arabic }
                 * { FEINIT \*charformat }{ CLIENTNO \* arabic \*charformat }{ MATTERNO \* arabic \*charformat }
                 */
                
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