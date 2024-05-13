namespace documorph.tests;

public class TableModelTests
{
    #region HideThisTableXml
    private const string TABLE_XML = @"
<w:tbl xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml"">
    <w:tblPr>
        <w:tblStyle w:val=""GridTable2""/>
        <w:tblW w:w=""0"" w:type=""auto""/>
        <w:tblLook w:val=""04A0"" w:firstRow=""1"" w:lastRow=""0"" w:firstColumn=""1"" w:lastColumn=""0"" w:noHBand=""0"" w:noVBand=""1""/>
    </w:tblPr>
    <w:tblGrid>
        <w:gridCol w:w=""1870""/>
        <w:gridCol w:w=""1870""/>
        <w:gridCol w:w=""1870""/>
        <w:gridCol w:w=""1870""/>
        <w:gridCol w:w=""1870""/>
    </w:tblGrid>
    <w:tr w:rsidR=""00FD6CC0"" w14:paraId=""0AE0F5B0"" w14:textId=""77777777"" w:rsidTr=""00FD6CC0"">
        <w:trPr>
            <w:cnfStyle w:val=""100000000000"" w:firstRow=""1"" w:lastRow=""0"" w:firstColumn=""0"" w:lastColumn=""0"" w:oddVBand=""0"" w:evenVBand=""0"" w:oddHBand=""0"" w:evenHBand=""0"" w:firstRowFirstColumn=""0"" w:firstRowLastColumn=""0"" w:lastRowFirstColumn=""0"" w:lastRowLastColumn=""0""/>
        </w:trPr>
        <w:tc>
            <w:tcPr>
                <w:cnfStyle w:val=""001000000000"" w:firstRow=""0"" w:lastRow=""0"" w:firstColumn=""1"" w:lastColumn=""0"" w:oddVBand=""0"" w:evenVBand=""0"" w:oddHBand=""0"" w:evenHBand=""0"" w:firstRowFirstColumn=""0"" w:firstRowLastColumn=""0"" w:lastRowFirstColumn=""0"" w:lastRowLastColumn=""0""/>
                <w:tcW w:w=""1870"" w:type=""dxa""/>
            </w:tcPr>
            <w:p w14:paraId=""60B09F80"" w14:textId=""1063A13E"" w:rsidR=""00FD6CC0"" w:rsidRDefault=""00FD6CC0"" w:rsidP=""00FD6CC0"">
                <w:r>
                    <w:t>Column 1</w:t>
                </w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:tcPr>
                <w:tcW w:w=""1870"" w:type=""dxa""/>
            </w:tcPr>
            <w:p w14:paraId=""72E3E1E4"" w14:textId=""3E7D3912"" w:rsidR=""00FD6CC0"" w:rsidRDefault=""00FD6CC0"" w:rsidP=""00FD6CC0"">
                <w:pPr>
                    <w:cnfStyle w:val=""100000000000"" w:firstRow=""1"" w:lastRow=""0"" w:firstColumn=""0"" w:lastColumn=""0"" w:oddVBand=""0"" w:evenVBand=""0"" w:oddHBand=""0"" w:evenHBand=""0"" w:firstRowFirstColumn=""0"" w:firstRowLastColumn=""0"" w:lastRowFirstColumn=""0"" w:lastRowLastColumn=""0""/>
                </w:pPr>
                <w:r>
                    <w:t>Column 2</w:t>
                </w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:tcPr>
                <w:tcW w:w=""1870"" w:type=""dxa""/>
            </w:tcPr>
            <w:p w14:paraId=""2B5DD35A"" w14:textId=""7CFD656A"" w:rsidR=""00FD6CC0"" w:rsidRDefault=""00FD6CC0"" w:rsidP=""00FD6CC0"">
                <w:pPr>
                    <w:cnfStyle w:val=""100000000000"" w:firstRow=""1"" w:lastRow=""0"" w:firstColumn=""0"" w:lastColumn=""0"" w:oddVBand=""0"" w:evenVBand=""0"" w:oddHBand=""0"" w:evenHBand=""0"" w:firstRowFirstColumn=""0"" w:firstRowLastColumn=""0"" w:lastRowFirstColumn=""0"" w:lastRowLastColumn=""0""/>
                </w:pPr>
                <w:r>
                    <w:t>Column 3</w:t>
                </w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:tcPr>
                <w:tcW w:w=""1870"" w:type=""dxa""/>
            </w:tcPr>
            <w:p w14:paraId=""4981F059"" w14:textId=""78730936"" w:rsidR=""00FD6CC0"" w:rsidRDefault=""00FD6CC0"" w:rsidP=""00FD6CC0"">
                <w:pPr>
                    <w:cnfStyle w:val=""100000000000"" w:firstRow=""1"" w:lastRow=""0"" w:firstColumn=""0"" w:lastColumn=""0"" w:oddVBand=""0"" w:evenVBand=""0"" w:oddHBand=""0"" w:evenHBand=""0"" w:firstRowFirstColumn=""0"" w:firstRowLastColumn=""0"" w:lastRowFirstColumn=""0"" w:lastRowLastColumn=""0""/>
                </w:pPr>
                <w:r>
                    <w:t>Column 4</w:t>
                </w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:tcPr>
                <w:tcW w:w=""1870"" w:type=""dxa""/>
            </w:tcPr>
            <w:p w14:paraId=""18EE5935"" w14:textId=""1979CADE"" w:rsidR=""00FD6CC0"" w:rsidRDefault=""00FD6CC0"" w:rsidP=""00FD6CC0"">
                <w:pPr>
                    <w:cnfStyle w:val=""100000000000"" w:firstRow=""1"" w:lastRow=""0"" w:firstColumn=""0"" w:lastColumn=""0"" w:oddVBand=""0"" w:evenVBand=""0"" w:oddHBand=""0"" w:evenHBand=""0"" w:firstRowFirstColumn=""0"" w:firstRowLastColumn=""0"" w:lastRowFirstColumn=""0"" w:lastRowLastColumn=""0""/>
                </w:pPr>
                <w:r>
                    <w:t>Column 5</w:t>
                </w:r>
            </w:p>
        </w:tc>
    </w:tr>
    <w:tr w:rsidR=""00FD6CC0"" w14:paraId=""55C6235E"" w14:textId=""77777777"" w:rsidTr=""00FD6CC0"">
        <w:trPr>
            <w:cnfStyle w:val=""000000100000"" w:firstRow=""0"" w:lastRow=""0"" w:firstColumn=""0"" w:lastColumn=""0"" w:oddVBand=""0"" w:evenVBand=""0"" w:oddHBand=""1"" w:evenHBand=""0"" w:firstRowFirstColumn=""0"" w:firstRowLastColumn=""0"" w:lastRowFirstColumn=""0"" w:lastRowLastColumn=""0""/>
        </w:trPr>
        <w:tc>
            <w:tcPr>
                <w:cnfStyle w:val=""001000000000"" w:firstRow=""0"" w:lastRow=""0"" w:firstColumn=""1"" w:lastColumn=""0"" w:oddVBand=""0"" w:evenVBand=""0"" w:oddHBand=""0"" w:evenHBand=""0"" w:firstRowFirstColumn=""0"" w:firstRowLastColumn=""0"" w:lastRowFirstColumn=""0"" w:lastRowLastColumn=""0""/>
                <w:tcW w:w=""1870"" w:type=""dxa""/>
            </w:tcPr>
            <w:p w14:paraId=""733E1396"" w14:textId=""4D26FD82"" w:rsidR=""00FD6CC0"" w:rsidRDefault=""00FD6CC0"" w:rsidP=""00FD6CC0"">
                <w:r>
                    <w:t>Cell 1 1</w:t>
                </w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:tcPr>
                <w:tcW w:w=""1870"" w:type=""dxa""/>
            </w:tcPr>
            <w:p w14:paraId=""2F82599B"" w14:textId=""5EB72AA3"" w:rsidR=""00FD6CC0"" w:rsidRDefault=""00FD6CC0"" w:rsidP=""00FD6CC0"">
                <w:pPr>
                    <w:cnfStyle w:val=""000000100000"" w:firstRow=""0"" w:lastRow=""0"" w:firstColumn=""0"" w:lastColumn=""0"" w:oddVBand=""0"" w:evenVBand=""0"" w:oddHBand=""1"" w:evenHBand=""0"" w:firstRowFirstColumn=""0"" w:firstRowLastColumn=""0"" w:lastRowFirstColumn=""0"" w:lastRowLastColumn=""0""/>
                </w:pPr>
                <w:r>
                    <w:t>Cell 2 1</w:t>
                </w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:tcPr>
                <w:tcW w:w=""1870"" w:type=""dxa""/>
            </w:tcPr>
            <w:p w14:paraId=""36B44961"" w14:textId=""43F5339A"" w:rsidR=""00FD6CC0"" w:rsidRDefault=""00FD6CC0"" w:rsidP=""00FD6CC0"">
                <w:pPr>
                    <w:cnfStyle w:val=""000000100000"" w:firstRow=""0"" w:lastRow=""0"" w:firstColumn=""0"" w:lastColumn=""0"" w:oddVBand=""0"" w:evenVBand=""0"" w:oddHBand=""1"" w:evenHBand=""0"" w:firstRowFirstColumn=""0"" w:firstRowLastColumn=""0"" w:lastRowFirstColumn=""0"" w:lastRowLastColumn=""0""/>
                </w:pPr>
                <w:r>
                    <w:t>Cell 3 1</w:t>
                </w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:tcPr>
                <w:tcW w:w=""1870"" w:type=""dxa""/>
            </w:tcPr>
            <w:p w14:paraId=""3D6FDF50"" w14:textId=""73E788BD"" w:rsidR=""00FD6CC0"" w:rsidRDefault=""00FD6CC0"" w:rsidP=""00FD6CC0"">
                <w:pPr>
                    <w:cnfStyle w:val=""000000100000"" w:firstRow=""0"" w:lastRow=""0"" w:firstColumn=""0"" w:lastColumn=""0"" w:oddVBand=""0"" w:evenVBand=""0"" w:oddHBand=""1"" w:evenHBand=""0"" w:firstRowFirstColumn=""0"" w:firstRowLastColumn=""0"" w:lastRowFirstColumn=""0"" w:lastRowLastColumn=""0""/>
                </w:pPr>
                <w:r>
                    <w:t>Cell 4 1</w:t>
                </w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:tcPr>
                <w:tcW w:w=""1870"" w:type=""dxa""/>
            </w:tcPr>
            <w:p w14:paraId=""07D6C813"" w14:textId=""2CD730CA"" w:rsidR=""00FD6CC0"" w:rsidRDefault=""00FD6CC0"" w:rsidP=""00FD6CC0"">
                <w:pPr>
                    <w:cnfStyle w:val=""000000100000"" w:firstRow=""0"" w:lastRow=""0"" w:firstColumn=""0"" w:lastColumn=""0"" w:oddVBand=""0"" w:evenVBand=""0"" w:oddHBand=""1"" w:evenHBand=""0"" w:firstRowFirstColumn=""0"" w:firstRowLastColumn=""0"" w:lastRowFirstColumn=""0"" w:lastRowLastColumn=""0""/>
                </w:pPr>
                <w:r>
                    <w:t>Cell 5 1</w:t>
                </w:r>
            </w:p>
        </w:tc>
    </w:tr>
</w:tbl>";
    #endregion
    
    [Fact]
    public void AppendMarkdown_Should_Append_Table_To_StringBuilder()
    {
        // Arrange
        var table = new Table(TABLE_XML);
        var hyperlinkRelationships = new List<HyperlinkRelationship>();
        var builder = new StringBuilder();

        // Act
        var tableModel = TableModel.FromTable(table, hyperlinkRelationships);

        // Assert
        Assert.Equal(2, tableModel.Rows.Count());
        Assert.Equal("Column 1", tableModel.Rows.First().First().Children.OfType<RunModel>().First().Text);
    }
}