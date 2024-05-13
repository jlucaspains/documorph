
namespace documorph.tests;

public class HyperlinkModelTests
{
    [Fact]
    public void AppendMarkdown_WhenHyperlinkIdIsNull_ShouldNotIncludeHyperlinkUrl()
    {
        // Arrange
        var hyperlink = new Hyperlink();
        var hyperlinkRelationships = new List<HyperlinkRelationship>();


        // Act
        var hyperlinkModel = HyperlinkModel.FromHyperlink(hyperlink, hyperlinkRelationships);

        // Assert
        Assert.DoesNotContain("http", hyperlinkModel.Uri);
    }

    [Fact]
    public void AppendMarkdown_WhenHyperlinkIdIsNotNull_ShouldIncludeHyperlinkUrl()
    {
        // Arrange
        var hyperlink = new Hyperlink(@"
<w:hyperlink 
xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" 
xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
r:id=""rId5"" w:history=""1"">
    <w:r w:rsidR=""00524EF3"" w:rsidRPr=""00524EF3"">
        <w:rPr>
            <w:rStyle w:val=""Hyperlink""/>
        </w:rPr>
        <w:t>A link</w:t>
    </w:r>
</w:hyperlink>");

        using var wordDoc = WordprocessingDocument.Create("AppendMarkdown_WhenHyperlinkIdIsNotNull_ShouldIncludeHyperlinkUrl.docx", WordprocessingDocumentType.Document);
        wordDoc.AddMainDocumentPart();
        var hyperlinkRel = wordDoc.MainDocumentPart!.AddHyperlinkRelationship(new Uri("https://example.com"), true, "rId5");

        // Act
        var hyperlinkModel = HyperlinkModel.FromHyperlink(hyperlink, [hyperlinkRel]);

        // Assert
        Assert.Equal("https://example.com/", hyperlinkModel.Uri);
        Assert.Equal("A link", hyperlinkModel.Runs.First().Text);
    }
}