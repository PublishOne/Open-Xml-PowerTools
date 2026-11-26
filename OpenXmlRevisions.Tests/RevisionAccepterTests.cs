// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using OpenXmlRevisions.Tests;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OxPt;

public class RaTests
{
    [Theory]
    [InlineData("RA001-Tracked-Revisions-01.docx")]
    [InlineData("RA001-Tracked-Revisions-02.docx")]

    public void RA001(string name)
    {
        DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
        FileInfo sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));
        var processedDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-processed-by-RevisionAccepter.docx")));

        using (WordprocessingDocument notAccepted = WordprocessingDocument.Open(sourceDocx.FullName, false))
        {
            using var fileStream = processedDestDocx.Create();
            using var afterAccepting = notAccepted.Clone(fileStream, true);
            RevisionAccepter.AcceptRevisions(afterAccepting);
        }
        using (WordprocessingDocument afterAccepting = WordprocessingDocument.Open(processedDestDocx.FullName, false))
        {
            Assert.False(RevisionAccepter.HasTrackedRevisions(afterAccepting));
            Assert.False(new TrackedChangesIdentifier().DescendantPartsOrSelfHasTrackedRevisions(afterAccepting.MainDocumentPart), "PublishOne still finds revisions");
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void AcceptContentControlWithFullyDeletedParagraph(bool hasSurvivingContent)
    {
        var deletedParagraph = new XElement(W.p,
            new XElement(W.pPr, new XElement(W.rPr,
                new XElement(W.del, new XAttribute(W.id, "1"), new XAttribute(W.author, "Author")))),
            new XElement(W.del, new XAttribute(W.id, "2"), new XAttribute(W.author, "Author"),
                new XElement(W.r, new XElement(W.delText, "Deleted"))));
        var content = new XElement(W.sdtContent,
            hasSurvivingContent ? new XElement(W.p, new XElement(W.r, new XElement(W.t, "Surviving"))) : null,
            deletedParagraph);
        var document = new XDocument(new XElement(W.document,
            new XElement(W.body,
                new XElement(W.sdt,
                    new XElement(W.sdtPr, new XElement(W.tag, new XAttribute(W.val, "Regression"))),
                    content),
                new XElement(W.sectPr))));

        using var stream = new MemoryStream();
        using var wordDocument = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var part = wordDocument.AddMainDocumentPart();
        part.PutXDocument(document);

        RevisionAccepter.AcceptRevisions(wordDocument);

        var accepted = part.GetXDocument();
        Assert.False(RevisionAccepter.HasTrackedRevisions(wordDocument));
        Assert.Equal(hasSurvivingContent ? "Surviving" : "", string.Concat(accepted.Descendants(W.t).Select(t => t.Value)));
        Assert.Empty(accepted.Descendants(W.delText));
        Assert.Equal(hasSurvivingContent ? 1 : 0, accepted.Descendants(W.sdt).Count());
        if (hasSurvivingContent)
        {
            var control = Assert.Single(accepted.Descendants(W.sdt));
            Assert.Equal("Regression", (string)control.Element(W.sdtPr).Element(W.tag).Attribute(W.val));
            Assert.Single(control.Element(W.sdtContent).Elements(W.p));
        }
    }

    [Theory]
    [InlineData("RevisionAccept-099.docx")]// (Skip ="This test is a known issue and needs to be investigated")
    [InlineData("RevisionAccept-100.docx")]

    public void RA099(string name)
    {
        DirectoryInfo sourceDir = new DirectoryInfo("../../../../AdditionalTestFiles/");
        FileInfo sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));
        var processedDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-processed-by-RevisionAccepter.docx")));

        using (WordprocessingDocument notAccepted = WordprocessingDocument.Open(sourceDocx.FullName, false))
        {
            using var fileStream = processedDestDocx.Create();
            using var afterAccepting = notAccepted.Clone(fileStream, true);
            RevisionAccepter.AcceptRevisions(afterAccepting);
        }
        using (WordprocessingDocument afterAccepting = WordprocessingDocument.Open(processedDestDocx.FullName, false))
        {
            Assert.False(RevisionAccepter.HasTrackedRevisions(afterAccepting));
            Assert.False(new TrackedChangesIdentifier().DescendantPartsOrSelfHasTrackedRevisions(afterAccepting.MainDocumentPart), "PublishOne still finds revisions");
        }
    }
}
