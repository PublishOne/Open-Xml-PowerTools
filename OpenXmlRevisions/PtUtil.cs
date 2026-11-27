// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections;
using System.Text;
using System.Xml.Linq;

namespace OpenXmlRevisions;

public static class PtExtensions
{
    public static string StringConcatenate(this IEnumerable<string> source)
    {
        return source.Aggregate(
            new StringBuilder(),
            (sb, s) => sb.Append(s),
            sb => sb.ToString());
    }

    public static string StringConcatenate<T>(this IEnumerable<T> source, Func<T, string> projectionFunc)
    {
        return source.Aggregate(
            new StringBuilder(),
            (sb, i) => sb.Append(projectionFunc(i)),
            sb => sb.ToString());
    }

    public static IEnumerable<IGrouping<TKey?, TSource>> GroupAdjacent<TSource, TKey>(
        this IEnumerable<TSource> source,
        Func<TSource, TKey?> keySelector)
    {
        var last = default(TKey?);
        var haveLast = false;
        var list = new List<TSource>();

        foreach (TSource s in source)
        {
            TKey? k = keySelector(s);
            if (haveLast)
            {
                if (!Equals(k, last))
                {
                    yield return new GroupOfAdjacent<TSource, TKey?>(list, last);

                    list = [s];
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
        {
            yield return new GroupOfAdjacent<TSource, TKey?>(list, last);
        }
    }

    private static void InitializeSiblingsReverseDocumentOrder(XElement element)
    {
        XElement? prev = null;
        foreach (XElement e in element.Elements())
        {
            e.AddAnnotation(new SiblingsReverseDocumentOrderInfo { PreviousSibling = prev });
            prev = e;
        }
    }

    public static IEnumerable<XElement> SiblingsBeforeSelfReverseDocumentOrder(
        this XElement element)
    {
        if (element.Annotation<SiblingsReverseDocumentOrderInfo>() == null)
        {
            InitializeSiblingsReverseDocumentOrder(element.Parent ?? throw new InvalidOperationException("Element has no parent"));
        }

        XElement current = element;
        while (true)
        {
            XElement? previousElement = current
                .Annotation<SiblingsReverseDocumentOrderInfo>()?
                .PreviousSibling;
            if (previousElement == null)
            {
                yield break;
            }

            yield return previousElement;

            current = previousElement;
        }
    }

    public static IEnumerable<XElement> DescendantsTrimmed(this XElement element,
        XName trimName)
    {
        return DescendantsTrimmed(element, e => e.Name == trimName);
    }

    public static IEnumerable<XElement> DescendantsTrimmed(this XElement element,
        Func<XElement, bool> predicate)
    {
        var iteratorStack = new Stack<IEnumerator<XElement>>();
        iteratorStack.Push(element.Elements().GetEnumerator());
        while (iteratorStack.Count > 0)
        {
            while (iteratorStack.Peek().MoveNext())
            {
                XElement currentXElement = iteratorStack.Peek().Current;
                if (predicate(currentXElement))
                {
                    yield return currentXElement;
                    continue;
                }
                yield return currentXElement;
                iteratorStack.Push(currentXElement.Elements().GetEnumerator());
            }
            iteratorStack.Pop();
        }
    }

    public static IEnumerable<TResult> Rollup<TSource, TResult>(
        this IEnumerable<TSource> source,
        TResult seed,
        Func<TSource, TResult, TResult> projection)
    {
        TResult nextSeed = seed;
        foreach (TSource src in source)
        {
            TResult projectedValue = projection(src, nextSeed);
            nextSeed = projectedValue;
            yield return projectedValue;
        }
    }

    public static IEnumerable<TResult> Rollup<TSource, TResult>(
        this IEnumerable<TSource> source,
        TResult seed,
        Func<TSource, TResult, int, TResult> projection)
    {
        TResult nextSeed = seed;
        int index = 0;
        foreach (TSource src in source)
        {
            TResult projectedValue = projection(src, nextSeed, index++);
            nextSeed = projectedValue;
            yield return projectedValue;
        }
    }

    public static IEnumerable<T> PtSkipLast<T>(this IEnumerable<T> source, int count)
    {
        var saveList = new Queue<T>();
        var saved = 0;
        foreach (T item in source)
        {
            if (saved < count)
            {
                saveList.Enqueue(item);
                ++saved;
                continue;
            }

            saveList.Enqueue(item);
            yield return saveList.Dequeue();
        }
    }
}

internal sealed class SiblingsReverseDocumentOrderInfo
{
    public required XElement? PreviousSibling { get; set; }
}

internal sealed class GroupOfAdjacent<TSource, TKey>(ICollection<TSource> source, TKey key) : IGrouping<TKey, TSource>
{
    public TKey Key { get; set; } = key;
    private readonly ICollection<TSource> _groupList = source ?? throw new ArgumentNullException(nameof(source));

    IEnumerator IEnumerable.GetEnumerator()
    {
        return ((IEnumerable<TSource>)this).GetEnumerator();
    }

    IEnumerator<TSource> IEnumerable<TSource>.GetEnumerator()
    {
        return _groupList.GetEnumerator();
    }
}
