namespace FFE
{
    public interface IFfeJsonParser : IFfeParser
    {
        string SelectByJsonPath(string jsonPath);

        string GetJson();
    }
}