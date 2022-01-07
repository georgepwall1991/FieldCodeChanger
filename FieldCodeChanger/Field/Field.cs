using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using FieldCodeChanger.ExtensionMethods;

namespace FieldCodeChanger.Field
{
    public static class Field
    {
        public enum FieldElementTypeEnum
        {
            Begin,
            InstrText,
            Separate,
            Result,
            End
        }

        public static string InstrText(XElement root, int id)
        {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            var relevantElements = root.Descendants()
                .Where(e =>
                {
                    var s = e.Annotation<Stack<FieldElementTypeInfo>>();
                    if (s != null)
                        return s.Any(z => z.Id == id &&
                                          z.FieldElementType == FieldElementTypeEnum.InstrText);
                    return false;
                })
                .ToList();
            var groupedSubFields = relevantElements
                .GroupAdjacent(e =>
                {
                    var s = e.Annotation<Stack<FieldElementTypeInfo>>();
                    var stackElement = s.FirstOrDefault(z => z.Id == id);
                    var elementsBefore = s.TakeWhile(z => z != stackElement);
                    return elementsBefore.Any();
                });
            var instrText = groupedSubFields
                .Select(g =>
                {
                    if (g.Key == false)
                        return g.Select(e =>
                            {
                                var s = e.Annotation<Stack<FieldElementTypeInfo>>();
                                var stackElement = s.FirstOrDefault(z => z.Id == id);
                                if (stackElement.FieldElementType == FieldElementTypeEnum.InstrText &&
                                    e.Name == w + "instrText")
                                    return e.Value;
                                return "";
                            })
                            .StringConcatenate();

                    {
                        var s = g.First().Annotation<Stack<FieldElementTypeInfo>>();
                        var stackElement = s.FirstOrDefault(z => z.Id == id);
                        var elementBefore = s.TakeWhile(z => z != stackElement).Last();
                        var subFieldId = elementBefore.Id;
                        return InstrText(root, subFieldId);
                    }
                })
                .StringConcatenate();
            return "{" + instrText + "}";
        }

        public static void AnnotateWithFieldInfo(OpenXmlPart part)
        {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            var root = part.GetXDocument().Root;
            var r = root.DescendantsAndSelf()
                .Rollup(
                    new FieldElementTypeStack
                    {
                        Id = 0,
                        FiStack = null
                    },
                    (e, s) =>
                    {
                        if (e.Name == w + "fldChar")
                        {
                            if (e.Attribute(w + "fldCharType").Value == "begin")
                            {
                                Stack<FieldElementTypeInfo> fis;
                                fis = s.FiStack == null
                                    ? new Stack<FieldElementTypeInfo>()
                                    : new Stack<FieldElementTypeInfo>(s.FiStack.Reverse());
                                fis.Push(
                                    new FieldElementTypeInfo
                                    {
                                        Id = s.Id + 1,
                                        FieldElementType = FieldElementTypeEnum.Begin
                                    });
                                return new FieldElementTypeStack
                                {
                                    Id = s.Id + 1,
                                    FiStack = fis
                                };
                            }

                            if (e.Attribute(w + "fldCharType")?.Value == "separate")
                            {
                                var fis = new Stack<FieldElementTypeInfo>(s.FiStack.Reverse());
                                var wfi = fis.Pop();
                                fis.Push(
                                    new FieldElementTypeInfo
                                    {
                                        Id = wfi.Id,
                                        FieldElementType = FieldElementTypeEnum.Separate
                                    });
                                return new FieldElementTypeStack
                                {
                                    Id = s.Id,
                                    FiStack = fis
                                };
                            }

                            if (e.Attribute(w + "fldCharType").Value == "end")
                            {
                                var fis = new Stack<FieldElementTypeInfo>(s.FiStack.Reverse());
                                var wfi = fis.Pop();
                                fis.Push(
                                    new FieldElementTypeInfo
                                    {
                                        Id = wfi.Id,
                                        FieldElementType = FieldElementTypeEnum.End
                                    });
                                return new FieldElementTypeStack
                                {
                                    Id = s.Id,
                                    FiStack = fis
                                };
                            }
                        }

                        if (s.FiStack == null || !s.FiStack.Any())
                            return s;
                        var wfi3 = s.FiStack.Peek();
                        switch (wfi3.FieldElementType)
                        {
                            case FieldElementTypeEnum.Begin:
                            {
                                var fis = new Stack<FieldElementTypeInfo>(s.FiStack.Reverse());
                                var wfi2 = fis.Pop();
                                fis.Push(
                                    new FieldElementTypeInfo
                                    {
                                        Id = wfi2.Id,
                                        FieldElementType = FieldElementTypeEnum.InstrText
                                    });
                                return new FieldElementTypeStack
                                {
                                    Id = s.Id,
                                    FiStack = fis
                                };
                            }
                            case FieldElementTypeEnum.Separate:
                            {
                                var fis = new Stack<FieldElementTypeInfo>(s.FiStack.Reverse());
                                var wfi2 = fis.Pop();
                                fis.Push(
                                    new FieldElementTypeInfo
                                    {
                                        Id = wfi2.Id,
                                        FieldElementType = FieldElementTypeEnum.Result
                                    });
                                return new FieldElementTypeStack
                                {
                                    Id = s.Id,
                                    FiStack = fis
                                };
                            }
                        }

                        if (wfi3.FieldElementType != FieldElementTypeEnum.End) return s;
                        {
                            var fis = new Stack<FieldElementTypeInfo>(s.FiStack.Reverse());
                            fis.Pop();
                            if (!fis.Any())
                                fis = null;
                            return new FieldElementTypeStack
                            {
                                Id = s.Id,
                                FiStack = fis
                            };
                        }
                    });
            var elementPlusInfo = root.DescendantsAndSelf().Zip(r, (t1, t2) => new
            {
                Element = t1,
                t2.Id,
                WmlFieldInfoStack = t2.FiStack
            });
            foreach (var item in elementPlusInfo)
                if (item.WmlFieldInfoStack != null)
                    item.Element.AddAnnotation(item.WmlFieldInfoStack);
        }

        private static string[] GetTokens(string field)
        {
            var state = State.InWhiteSpace;
            var tokenStart = 0;
            var quoteStart = char.MinValue;
            var tokens = new List<string>();
            for (var c = 0; c < field.Length; c++)
            {
                if (char.IsWhiteSpace(field[c]))
                {
                    switch (state)
                    {
                        case State.InToken:
                            tokens.Add(field.Substring(tokenStart, c - tokenStart));
                            state = State.InWhiteSpace;
                            continue;
                        case State.OnOpeningQuote:
                            tokenStart = c;
                            state = State.InQuotedToken;
                            break;
                    }

                    if (state == State.OnClosingQuote)
                        state = State.InWhiteSpace;
                    continue;
                }

                if (field[c] == '\\')
                    if (state == State.InQuotedToken)
                    {
                        state = State.OnBackslash;
                        continue;
                    }

                if (state == State.OnBackslash)
                {
                    state = State.InQuotedToken;
                    continue;
                }

                if (field[c] == '"' || field[c] == '\'' || field[c] == 0x201d)
                {
                    switch (state)
                    {
                        case State.InWhiteSpace:
                            quoteStart = field[c];
                            state = State.OnOpeningQuote;
                            continue;
                        case State.InQuotedToken:
                        {
                            if (field[c] == quoteStart)
                            {
                                tokens.Add(field.Substring(tokenStart, c - tokenStart));
                                state = State.OnClosingQuote;
                            }

                            continue;
                        }
                        case State.OnOpeningQuote when field[c] == quoteStart:
                            state = State.OnClosingQuote;
                            continue;
                        case State.OnOpeningQuote:
                            tokenStart = c;
                            state = State.InQuotedToken;
                            break;
                    }

                    continue;
                }

                switch (state)
                {
                    case State.InWhiteSpace:
                        tokenStart = c;
                        state = State.InToken;
                        continue;
                    case State.OnOpeningQuote:
                        tokenStart = c;
                        state = State.InQuotedToken;
                        continue;
                    case State.OnClosingQuote:
                        tokenStart = c;
                        state = State.InToken;
                        break;
                }
            }

            if (state == State.InToken)
                tokens.Add(field.Substring(tokenStart, field.Length - tokenStart));
            return tokens.ToArray();
        }

        public static FieldInfo ParseField(string field)
        {
            var emptyField = new FieldInfo
            {
                FieldType = "",
                Arguments = new string[] { },
                Switches = new string[] { }
            };

            if (field.Length == 0)
                return emptyField;
            var fieldType = field.TrimStart().Split(' ').FirstOrDefault();
            if (fieldType == null || fieldType.ToUpper() != "HYPERLINK")
                return emptyField;
            var tokens = GetTokens(field);
            if (tokens.Length == 0)
                return emptyField;
            var fieldInfo = new FieldInfo
            {
                FieldType = tokens[0],
                Switches = tokens.Where(t => t[0] == '\\').ToArray(),
                Arguments = tokens.Skip(1).Where(t => t[0] != '\\').ToArray()
            };
            return fieldInfo;
        }

        private enum State
        {
            InToken,
            InWhiteSpace,
            InQuotedToken,
            OnOpeningQuote,
            OnClosingQuote,
            OnBackslash
        }

        public class FieldInfo
        {
            public string[] Arguments;
            public string FieldType;
            public string[] Switches;
        }

        public class FieldElementTypeInfo
        {
            public FieldElementTypeEnum FieldElementType;
            public int Id;
        }

        private class FieldElementTypeStack
        {
            public Stack<FieldElementTypeInfo> FiStack;
            public int Id;
        }
    }
}