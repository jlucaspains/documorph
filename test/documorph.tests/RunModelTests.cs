using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace documorph.tests;

public class RunModelTests
{
    [Fact]
    public void AppendMarkdown_Should_AppendBasicText()
    {
        // Arrange
        var run = new Run(@"
<w:r xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:t xml:space=""preserve"">A simple text without formatting</w:t>
</w:r>");
        var parts = new List<IdPartPair>();
        var model = new RunModel(run, parts);
        var builder = new StringBuilder();

        // Act
        model.AppendMarkdown(builder);

        // Assert
        Assert.Equal("A simple text without formatting", builder.ToString());
    }

    [Fact]
    public void AppendMarkdown_Should_AppendBoldText()
    {
        // Arrange
        var run = new Run(@"
<w:r w:rsidRPr=""00524EF3"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:rPr>
        <w:b/>
        <w:bCs/>
    </w:rPr>
    <w:t>A bold text</w:t>
</w:r>");
        var parts = new List<DocumentFormat.OpenXml.Packaging.IdPartPair>();
        var model = new RunModel(run, parts);
        var builder = new StringBuilder();

        // Act
        model.AppendMarkdown(builder);

        // Assert
        Assert.Equal("**A bold text**", builder.ToString());
    }

    [Fact]
    public void AppendMarkdown_Should_AppendItalicText()
    {
        // Arrange
        var run = new Run(@"
<w:r w:rsidRPr=""00524EF3"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:rPr>
        <w:i/>
        <w:iCs/>
    </w:rPr>
    <w:t>An italic text</w:t>
</w:r>");
        var parts = new List<IdPartPair>();
        var model = new RunModel(run, parts);
        var builder = new StringBuilder();

        // Act
        model.AppendMarkdown(builder);

        // Assert
        Assert.Equal("*An italic text*", builder.ToString());
    }

    [Fact]
    public void AppendMarkdown_Should_AppendStrikedText()
    {
        // Arrange
        var run = new Run(@"
<w:r w:rsidRPr=""00B448D2"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:rPr>
        <w:strike/>
    </w:rPr>
    <w:t>Striked</w:t>
</w:r>
");
        var parts = new List<IdPartPair>();
        var model = new RunModel(run, parts);
        var builder = new StringBuilder();

        // Act
        model.AppendMarkdown(builder);

        // Assert
        Assert.Equal("~~Striked~~", builder.ToString());
    }

    [Fact]
    public void AppendMarkdown_Should_AppendUnderlineText()
    {
        // Arrange
        var run = new Run(@"
<w:r w:rsidRPr=""00B448D2"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:rPr>
        <w:u w:val=""single""/>
    </w:rPr>
    <w:t>Underline</w:t>
</w:r>
");
        var parts = new List<IdPartPair>();
        var model = new RunModel(run, parts);
        var builder = new StringBuilder();

        // Act
        model.AppendMarkdown(builder);

        // Assert
        Assert.Equal("__Underline__", builder.ToString());
    }

    [Fact]
    public void AppendMarkdown_Should_AppendImage()
    {
        // Arrange
        var run = new Run(@"
<w:r xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing""
xmlns:wp14=""http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing""
xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
    <w:rPr>
        <w:noProof/>
    </w:rPr>
    <w:lastRenderedPageBreak/>
    <w:drawing>
        <wp:inline distT=""0"" distB=""0"" distL=""0"" distR=""0"" wp14:anchorId=""3687BF73"" wp14:editId=""425059AE"">
            <wp:extent cx=""5943600"" cy=""3777615""/>
            <wp:effectExtent l=""0"" t=""0"" r=""0"" b=""0""/>
            <wp:docPr id=""1012819237"" name=""Picture 1"" descr=""A screenshot of a computer Description automatically generated""/>
            <wp:cNvGraphicFramePr>
                <a:graphicFrameLocks xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" noChangeAspect=""1""/>
            </wp:cNvGraphicFramePr>
            <a:graphic xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
                <a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
                    <pic:pic xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
                        <pic:nvPicPr>
                            <pic:cNvPr id=""1012819237"" name=""Picture 1"" descr=""A screenshot of a computer Description automatically generated""/>
                            <pic:cNvPicPr/>
                        </pic:nvPicPr>
                        <pic:blipFill>
                            <a:blip r:embed=""rId6"" cstate=""print"">
                                <a:extLst>
                                    <a:ext uri=""{28A0092B-C50C-407E-A947-70E740481C1C}"">
                                        <a14:useLocalDpi xmlns:a14=""http://schemas.microsoft.com/office/drawing/2010/main"" val=""0""/>
                                    </a:ext>
                                </a:extLst>
                            </a:blip>
                            <a:stretch>
                                <a:fillRect/>
                            </a:stretch>
                        </pic:blipFill>
                        <pic:spPr>
                            <a:xfrm>
                                <a:off x=""0"" y=""0""/>
                                <a:ext cx=""5943600"" cy=""3777615""/>
                            </a:xfrm>
                            <a:prstGeom prst=""rect"">
                                <a:avLst/>
                            </a:prstGeom>
                        </pic:spPr>
                    </pic:pic>
                </a:graphicData>
            </a:graphic>
        </wp:inline>
    </w:drawing>
</w:r>");
        using var wordDoc = WordprocessingDocument.Create("test.docx", WordprocessingDocumentType.Document);
        wordDoc.AddMainDocumentPart();
        ImagePart imagePart = wordDoc.MainDocumentPart!.AddImagePart(ImagePartType.Jpeg);

        var parts = new List<IdPartPair>{
            new("rId6", imagePart)
        };
        var model = new RunModel(run, parts);
        var builder = new StringBuilder();

        // Act
        model.AppendMarkdown(builder);

        // Assert
        // Add your assertions here
        Assert.Equal("![A screenshot of a computer Description automatically generated](image.jpg)", builder.ToString());
    }

    [Fact]
    public void AppendMarkdown_Should_AppendBreak()
    {
        // Arrange
        var run = new Run(@"
<w:r xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:br w:type=""page""/>
</w:r>");
        var parts = new List<IdPartPair>();
        var model = new RunModel(run, parts);
        var builder = new StringBuilder();

        // Act
        model.AppendMarkdown(builder);

        // Assert
        Assert.Equal(Environment.NewLine, builder.ToString());
    }
}