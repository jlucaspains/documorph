namespace documorph.tests;

public class ProgramTests
{
    [Fact]
    public void NoParametersReturn1()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [Array.Empty<string>()]);

        Assert.Equal(1, result);
    }

    [Fact]
    public void MissingCommandReturn1()
    {
        var entryPoint = typeof(Program).Assembly.EntryPoint!;
        var result = entryPoint.Invoke(null, [new string[] { }]);

        Assert.Equal(1, result);
    }
}