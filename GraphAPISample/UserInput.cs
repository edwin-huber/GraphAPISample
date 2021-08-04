using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GraphAPISample
{
    internal static class UserInput
    {
        internal static bool GetUserYesNo(string prompt)
        {
            Console.Write($"{prompt} (y/n)");
            ConsoleKeyInfo confirm;
            do
            {
                confirm = Console.ReadKey(true);
            }
            while (confirm.Key != ConsoleKey.Y && confirm.Key != ConsoleKey.N);

            Console.WriteLine();
            return (confirm.Key == ConsoleKey.Y);
        }

        internal static string GetUserInput(
            string fieldName,
            bool isRequired,
            Func<string, bool> validate)
        {
            string returnValue = null;
            do
            {
                Console.Write($"Enter a {fieldName}: ");
                if (!isRequired)
                {
                    Console.Write("(ENTER to skip) ");
                }
                var input = Console.ReadLine();

                if (string.IsNullOrEmpty(input)) continue;
                if (validate.Invoke(input))
                {
                    returnValue = input;
                }
            }
            while (string.IsNullOrEmpty(returnValue) && isRequired);

            return returnValue;
        }
    }
}
