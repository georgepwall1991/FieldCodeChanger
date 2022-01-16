using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace FieldCodeChanger.ExtensionMethods;

public static class Extensions
{
    public static XDocument GetXDocument(this OpenXmlPart part)
    {
        var partXDocument = part.Annotation<XDocument>();
        if (partXDocument != null)
            return partXDocument;
        using (var partStream = part.GetStream())
        using (var partXmlReader = XmlReader.Create(partStream))
        {
            partXDocument = XDocument.Load(partXmlReader);
        }

        part.AddAnnotation(partXDocument);
        return partXDocument;
    }

    public static IEnumerable<IGrouping<TKey, TSource>> GroupAdjacent<TSource, TKey>(
        this IEnumerable<TSource> source,
        Func<TSource, TKey> keySelector)
    {
        var last = default(TKey);
        var haveLast = false;
        var list = new List<TSource>();

        foreach (var s in source)
        {
            var k = keySelector(s);
            if (haveLast)
            {
                if (!k.Equals(last))
                {
                    yield return new GroupOfAdjacent<TSource, TKey>(list, last);
                    list = new List<TSource> { s };
                    last = k;
                }
                else
                {
                    list.Add(s);
                    last = k;
                }
            }
            else
            {
                list.Add(s);
                last = k;
                haveLast = true;
            }
        }

        if (haveLast)
            yield return new GroupOfAdjacent<TSource, TKey>(list, last);
    }

    public static string StringConcatenate(this IEnumerable<string> source)
    {
        var sb = new StringBuilder();
        foreach (var s in source)
            sb.Append(s);
        return sb.ToString();
    }

    public static IEnumerable<TResult> Rollup<TSource, TResult>(
        this IEnumerable<TSource> source,
        TResult seed,
        Func<TSource, TResult, TResult> projection)
    {
        var nextSeed = seed;
        foreach (var src in source)
        {
            var projectedValue = projection(src, nextSeed);
            nextSeed = projectedValue;
            yield return projectedValue;
        }
    }
}

public class GroupOfAdjacent<TSource, TKey> : IGrouping<TKey, TSource>
{
    public GroupOfAdjacent(List<TSource> source, TKey key)
    {
        GroupList = source;
        Key = key;
    }

    private List<TSource> GroupList { get; }
    public TKey Key { get; }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return ((IEnumerable<TSource>)this).GetEnumerator();
    }

    IEnumerator<TSource>
        IEnumerable<TSource>.GetEnumerator()
    {
        return ((IEnumerable<TSource>)GroupList).GetEnumerator();
    }
}