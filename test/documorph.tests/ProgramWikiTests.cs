namespace documorph.tests;

public class ProgramWikiTests
{
    [Fact]
    public void MissingInReturn1()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "wiki", "--out", "test.wiki" }]);

        Assert.Equal(1, result);
    }

    [Fact]
    public void MissingOutReturn1()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "wiki", "--in", "test.docx" }]);

        Assert.Equal(1, result);
    }

    [Fact]
    public void WillConvertSimpleDocx()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "wiki", "--in", "./test_data/simpletest.docx", "--out", "./test_data/WillConvertSimpleDocx.wiki" }]);

        Assert.Equal(0, result);
        Assert.True(File.Exists("./test_data/WillConvertSimpleDocx.wiki"));
    }
    
    [Fact]
    public void ValidParameters()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "wiki", "--in", "./test_data/docconverter.docx", "--out", "./test_data/Test_ValidParameters.wiki" }]);

        Assert.Equal(0, result);
        Assert.True(File.Exists("./test_data/Test_ValidParameters.wiki"));
    }
    
    [Fact]
    public void WillCreateOutputDirectoryIfNeeded()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "wiki", "--in", "./test_data/docconverter.docx", "--out", "./test_data2/Test_ValidParameters.wiki" }]);

        Assert.Equal(0, result);
        Assert.True(File.Exists("./test_data2/Test_ValidParameters.wiki"));
    }
    
    [Fact]
    public void InvalidInFileParameters()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "wiki", "--in", "test_data/docconverterbad.docx", "--out", "test_data/Test_ValidParameters.wiki" }]);

        Assert.Equal(1, result);
    }
    
    [Fact]
    public void InvalidOutFileParameters()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "wiki", "--in", "test_data/docconverter.docx", "--out", "" }]);

        Assert.Equal(1, result);
    }
    
    [Fact]
    public void WillFailIfInputIsDirectoryAndOutputIsNotDirectory()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "wiki", "--in", "test_data/", "--out", "test_data/docconverter.docx" }]);

        Assert.Equal(1, result);
    }

    [Fact]
    public void WillConvertAnEntireFolder()
    {
        Directory.CreateDirectory("WillConvertAnEntireFolderwiki/");
        File.Copy("test_data/docconverter.docx", "WillConvertAnEntireFolderwiki/docconverter1.docx", true);
        File.Copy("test_data/docconverter.docx", "WillConvertAnEntireFolderwiki/docconverter2.docx", true);
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { "wiki", "--in", "WillConvertAnEntireFolderwiki/", "--out", "WillConvertAnEntireFolderwiki/" }]);

        Assert.Equal(0, result);
        Assert.Equal(Directory.GetFiles("WillConvertAnEntireFolderwiki/", "*.docx").Length, 
            Directory.GetFiles("WillConvertAnEntireFolderwiki/", "*.wiki").Length);
    }
}