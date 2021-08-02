using System;

namespace ConsoleApp
{
    public class UserConsoleInput : IUserInput
    {
        public string GetFromPath()
        {
            return Console.ReadLine().Trim();
        }
        public string GetInPath()
        {
            return Console.ReadLine().Trim();
        }
    }
}
