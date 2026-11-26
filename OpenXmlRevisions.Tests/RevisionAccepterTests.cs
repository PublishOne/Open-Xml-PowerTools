// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using OpenXmlRevisions.Tests;
using System.IO;
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
    // [InlineData("RevisionAccept-099.docx")] (Skip ="This test is a known issue and needs to be investigated")
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
