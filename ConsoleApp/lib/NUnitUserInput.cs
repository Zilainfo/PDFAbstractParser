namespace ConsoleApp
{
    public class NUnitUserInput : IUserInput
    {
        public string fromPath { get; set; }
        public string InPath { get; set; }
        public string GetFromPath()
        {
            return fromPath;
        }
        public string GetInPath()
        {
            return InPath;
        }
    }
}
