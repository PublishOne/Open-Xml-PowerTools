// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;
using System.Xml;
using System.Xml.Linq;

// ReSharper disable InconsistentNaming

namespace OpenXmlRevisions;

public class PowerToolsInvalidDataException : Exception
{
    public PowerToolsInvalidDataException(string message) : base(message) { }

    public PowerToolsInvalidDataException()
    {
    }

    public PowerToolsInvalidDataException(string message, Exception innerException) : base(message, innerException)
    {
    }
}

public static class PtOpenXmlExtensions
{
    public static XDocument GetXDocument(this OpenXmlPart part)
    {
        ArgumentNullException.ThrowIfNull(part);

        var partXDocument = part.Annotation<XDocument>();
        if (partXDocument != null)
        {
            return partXDocument;
        }

        using (Stream partStream = part.GetStream())
        {
            if (partStream.Length == 0)
            {
                partXDocument = new XDocument
                {
                    Declaration = new XDeclaration("1.0", "UTF-8", "yes")
                };
            }
            else
            {
                using (var partXmlReader = XmlReader.Create(partStream))
                {
                    partXDocument = XDocument.Load(partXmlReader);
                }
            }
        }

        part.AddAnnotation(partXDocument);
        return partXDocument;
    }

    public static void PutXDocument(this OpenXmlPart part)
    {
        ArgumentNullException.ThrowIfNull(part);

        XDocument partXDocument = part.GetXDocument();
        if (partXDocument != null)
        {
            using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
            using (var partXmlWriter = XmlWriter.Create(partStream))
            {
                partXDocument.Save(partXmlWriter);
            }
        }
    }

    public static void PutXDocument(this OpenXmlPart part, XDocument document)
    {
        ArgumentNullException.ThrowIfNull(part);
        ArgumentNullException.ThrowIfNull(document);

        using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
        using (var partXmlWriter = XmlWriter.Create(partStream))
        {
            document.Save(partXmlWriter);
        }

        part.RemoveAnnotations<XDocument>();
        part.AddAnnotation(document);
    }

    public static IEnumerable<XElement> LogicalChildrenContent(this XElement element)
    {
        ArgumentNullException.ThrowIfNull(element);
        if (element.Name == W.document)
        {
            return element.Descendants(W.body).Take(1);
        }

        if (element.Name == W.body ||
            element.Name == W.tc ||
            element.Name == W.txbxContent)
        {
            return element
                .DescendantsTrimmed(e =>
                    e.Name == W.tbl ||
                    e.Name == W.p)
                .Where(e =>
                    e.Name == W.p ||
                    e.Name == W.tbl);
        }

        if (element.Name == W.tbl)
        {
            return element
                .DescendantsTrimmed(W.tr)
                .Where(e => e.Name == W.tr);
        }

        if (element.Name == W.tr)
        {
            return element
                .DescendantsTrimmed(W.tc)
                .Where(e => e.Name == W.tc);
        }

        if (element.Name == W.p)
        {
            return element
                .DescendantsTrimmed(e => e.Name == W.r ||
                    e.Name == W.pict ||
                    e.Name == W.drawing)
                .Where(e => e.Name == W.r ||
                    e.Name == W.pict ||
                    e.Name == W.drawing);
        }

        if (element.Name == W.r)
        {
            return element
                .DescendantsTrimmed(e => W.SubRunLevelContent.Contains(e.Name))
                .Where(e => W.SubRunLevelContent.Contains(e.Name));
        }

        if (element.Name == MC.AlternateContent)
        {
            return element
                .DescendantsTrimmed(e =>
                    e.Name == W.pict ||
                    e.Name == W.drawing ||
                    e.Name == MC.Fallback)
                .Where(e =>
                    e.Name == W.pict ||
                    e.Name == W.drawing);
        }

        if (element.Name == W.pict || element.Name == W.drawing)
        {
            return element
                .DescendantsTrimmed(W.txbxContent)
                .Where(e => e.Name == W.txbxContent);
        }

        return XElement.EmptySequence;
    }

    public static IEnumerable<XElement> LogicalChildrenContent(this XElement element, XName name)
    {
        return element.LogicalChildrenContent().Where(e => e.Name == name);
    }

    public static IEnumerable<OpenXmlPart> ContentParts(this WordprocessingDocument doc)
    {
        ArgumentNullException.ThrowIfNull(doc);
        if (doc.MainDocumentPart is null)
        {
            yield break;
        }
        yield return doc.MainDocumentPart;

        foreach (var hdr in doc.MainDocumentPart.HeaderParts)
        {
            yield return hdr;
        }

        foreach (var ftr in doc.MainDocumentPart.FooterParts)
        {
            yield return ftr;
        }

        if (doc.MainDocumentPart.FootnotesPart != null)
        {
            yield return doc.MainDocumentPart.FootnotesPart;
        }

        if (doc.MainDocumentPart.EndnotesPart != null)
        {
            yield return doc.MainDocumentPart.EndnotesPart;
        }
    }

    /// <summary>
    /// Creates a complete list of all parts contained in the <see cref="OpenXmlPartContainer"/>.
    /// </summary>
    /// <param name="container">
    /// A <see cref="WordprocessingDocument"/>.
    /// </param>
    /// <returns>list of <see cref="OpenXmlPart"/>s contained in the <see cref="OpenXmlPartContainer"/>.</returns>
    public static ICollection<OpenXmlPart> GetAllParts(this OpenXmlPartContainer container)
    {
        ArgumentNullException.ThrowIfNull(container);
        // Use a HashSet so that parts are processed only once.
        var partList = new HashSet<OpenXmlPart>();

        foreach (IdPartPair p in container.Parts)
        {
            AddPart(partList, p.OpenXmlPart);
        }

        return partList.OrderBy(p => p.ContentType).ThenBy(p => p.Uri.ToString()).ToList();
    }

    private static void AddPart(HashSet<OpenXmlPart> partList, OpenXmlPart part)
    {
        if (partList.Contains(part))
        {
            return;
        }

        partList.Add(part);
        foreach (IdPartPair p in part.Parts)
        {
            AddPart(partList, p.OpenXmlPart);
        }
    }
}

public sealed class FieldInfo
{
    public required IReadOnlyList<string> Switches { get; init; }
    public required IReadOnlyList<string> Arguments { get; init; }

    public required string FieldType { get; init; }
}

public static class FieldParser
{
    private enum State
    {
        InToken,
        InWhiteSpace,
        InQuotedToken,
        OnOpeningQuote,
        OnClosingQuote,
        OnBackslash,
    }

    private static List<string> GetTokens(string field)
    {
        State state = State.InWhiteSpace;
        int tokenStart = 0;
        char quoteStart = char.MinValue;
        var tokens = new List<string>();
        for (int c = 0; c < field.Length; c++)
        {
            if (char.IsWhiteSpace(field[c]))
            {
                if (state == State.InToken)
                {
                    tokens.Add(field.Substring(tokenStart, c - tokenStart));
                    state = State.InWhiteSpace;
                    continue;
                }
                if (state == State.OnOpeningQuote)
                {
                    tokenStart = c;
                    state = State.InQuotedToken;
                }
                if (state == State.OnClosingQuote)
                {
                    state = State.InWhiteSpace;
                }

                continue;
            }
            if (field[c] == '\\')
            {
                if (state == State.InQuotedToken)
                {
                    state = State.OnBackslash;
                    continue;
                }
            }
            if (state == State.OnBackslash)
            {
                state = State.InQuotedToken;
                continue;
            }
            if (field[c] is '"' or '\'' or (char)0x201d)
            {
                if (state == State.InWhiteSpace)
                {
                    quoteStart = field[c];
                    state = State.OnOpeningQuote;
                    continue;
                }
                if (state == State.InQuotedToken)
                {
                    if (field[c] == quoteStart)
                    {
                        tokens.Add(field.Substring(tokenStart, c - tokenStart));
                        state = State.OnClosingQuote;
                        continue;
                    }
                    continue;
                }
                if (state == State.OnOpeningQuote)
                {
                    if (field[c] == quoteStart)
                    {
                        state = State.OnClosingQuote;
                        continue;
                    }
                    else
                    {
                        tokenStart = c;
                        state = State.InQuotedToken;
                        continue;
                    }
                }
                continue;
            }
            if (state == State.InWhiteSpace)
            {
                tokenStart = c;
                state = State.InToken;
                continue;
            }
            if (state == State.OnOpeningQuote)
            {
                tokenStart = c;
                state = State.InQuotedToken;
                continue;
            }
            if (state == State.OnClosingQuote)
            {
                tokenStart = c;
                state = State.InToken;
                continue;
            }
        }
        if (state == State.InToken)
        {
            tokens.Add(field[tokenStart..]);
        }

        return tokens;
    }

    public static FieldInfo ParseField(string field)
    {
        ArgumentException.ThrowIfNullOrEmpty(field);
        var emptyField = new FieldInfo
        {
            FieldType = "",
            Arguments = [],
            Switches = [],
        };

        if (field.Length == 0)
        {
            return emptyField;
        }
        var spanField = field.AsSpan().TrimStart();
        int spaceIndex = spanField.IndexOf(' ');
        var fieldTypeSpan = spaceIndex == -1 ? spanField : spanField.Slice(0, spaceIndex);
        if (fieldTypeSpan.IsEmpty || !fieldTypeSpan.Equals("HYPERLINK", StringComparison.OrdinalIgnoreCase) && !fieldTypeSpan.Equals("REF", StringComparison.OrdinalIgnoreCase))
        {
            return emptyField;
        }

        List<string> tokens = GetTokens(field);
        if (tokens.Count == 0)
        {
            return emptyField;
        }

        var fieldInfo = new FieldInfo()
        {
            FieldType = tokens[0],
            Switches = [.. tokens.Where(t => t[0] == '\\')],
            Arguments = [.. tokens.Skip(1).Where(t => t[0] != '\\')],
        };
        return fieldInfo;
    }
}

public static class PtOpenXml
{
    public static readonly XNamespace ptOpenXml = "http://powertools.codeplex.com/documentbuilder/2011/insert";
    public static readonly XName Insert = ptOpenXml + "Insert";
    public static readonly XName Id = "Id";

    public static readonly XNamespace pt = "http://powertools.codeplex.com/2011";
    public static readonly XName Uri = pt + "Uri";
    public static readonly XName Unid = pt + "Unid";
    public static readonly XName SHA1Hash = pt + "SHA1Hash";
    public static readonly XName CorrelatedSHA1Hash = pt + "CorrelatedSHA1Hash";
    public static readonly XName StructureSHA1Hash = pt + "StructureSHA1Hash";
    public static readonly XName CorrelationSet = pt + "CorrelationSet";
    public static readonly XName Status = pt + "Status";

    public static readonly XName Level = pt + "Level";
    public static readonly XName IndentLevel = pt + "IndentLevel";
    public static readonly XName ContentType = pt + "ContentType";

    public static readonly XName trPr = pt + "trPr";
    public static readonly XName tcPr = pt + "tcPr";
    public static readonly XName rPr = pt + "rPr";
    public static readonly XName pPr = pt + "pPr";
    public static readonly XName tblPr = pt + "tblPr";
    public static readonly XName style = pt + "style";

    public static readonly XName FontName = pt + "FontName";
    public static readonly XName LanguageType = pt + "LanguageType";
    public static readonly XName AbstractNumId = pt + "AbstractNumId";
    public static readonly XName HtmlStructure = pt + "HtmlStructure";
    public static readonly XName HtmlStyle = pt + "HtmlStyle";
    public static readonly XName StyleName = pt + "StyleName";
    public static readonly XName TabWidth = pt + "TabWidth";
    public static readonly XName Leader = pt + "Leader";

    public static readonly XName ListItemRun = pt + "ListItemRun";

    public static readonly XName HtmlToWmlCssWidth = pt + "HtmlToWmlCssWidth";
}

public class InvalidOpenXmlDocumentException : Exception
{
    public InvalidOpenXmlDocumentException(string message) : base(message) { }

    public InvalidOpenXmlDocumentException()
    {
    }

    public InvalidOpenXmlDocumentException(string message, Exception innerException) : base(message, innerException)
    {
    }
}

public class OpenXmlPowerToolsException : Exception
{
    public OpenXmlPowerToolsException(string message) : base(message) { }

    public OpenXmlPowerToolsException()
    {
    }

    public OpenXmlPowerToolsException(string message, Exception innerException) : base(message, innerException)
    {
    }
}

public class ColumnReferenceOutOfRangeException : Exception
{
    public ColumnReferenceOutOfRangeException(string columnReference)
        : base($"Column reference ({columnReference}) is out of range.")
    {
    }

    public ColumnReferenceOutOfRangeException()
    {
    }

    public ColumnReferenceOutOfRangeException(string message, Exception innerException) : base(message, innerException)
    {
    }
}

public class WorksheetAlreadyExistsException : Exception
{
    public WorksheetAlreadyExistsException(string sheetName)
        : base($"The worksheet ({sheetName}) already exists.")
    {
    }

    public WorksheetAlreadyExistsException()
    {
    }

    public WorksheetAlreadyExistsException(string message, Exception innerException) : base(message, innerException)
    {
    }
}
