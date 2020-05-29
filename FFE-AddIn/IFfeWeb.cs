namespace FFE
{
    interface IFfeWeb
    {
        dynamic Load(string url);

        string SelectByXPath(string xPath);

        string SelectByCssSelector(string cssSelector);
    }
}
