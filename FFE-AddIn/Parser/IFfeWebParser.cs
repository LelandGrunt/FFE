namespace FFE
{
    public interface IFfeWebParser : IFfeParser
    {
        string SelectByXPath(string xPath);

        string SelectByCssSelector(string cssSelector);

        string GetHtml();
    }
}