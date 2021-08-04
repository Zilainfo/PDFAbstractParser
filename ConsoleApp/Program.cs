using System;


namespace ConsoleApp
{

    class Program
    {
        static void Main(string[] args)
        {
            /*      NUnitUserInput input = new NUnitUserInput();
                  input.fromPath = @"C:\Users\kstas\Desktop\Work\Abstract Book from the 5th World Psoriasis and Psoriatic Arthritis Conference 2018.pdf";
                  input.InPath = @"C:\Users\kstas\Desktop\Work\Data Entry - 5th World Psoriasis & Psoriatic Arthritis Conference 2018 - Case format.xlsx";*/
            UserConsoleInput input = new UserConsoleInput();
            Parser parser = new Parser(input, 44);

            parser.GetText();
            Console.WriteLine("Finish");
        }
    }

}
