using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System.Security;

namespace ConsoleApp1
{
    class Password
    {
        /// <summary>
        /// Method used to take the Password from the User
        /// </summary>
        /// <returns></returns>
        public string EnterPassword()
        {
            Console.WriteLine();
            Console.WriteLine("Enter Password");
            ConsoleKeyInfo key;
            String password = "";
            do
            {
                key = Console.ReadKey(true);

                if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                {
                    password += key.KeyChar;
                    Console.Write("*");

                }
                else
                {
                    if (key.Key == ConsoleKey.Backspace && password.Length > 0)
                    {

                        password = password.Substring(0, password.Length - 1);
                        Console.Write("\b \b");
                    }

                }
            } while (key.Key != ConsoleKey.Enter);
            return password;
        }

        /// <summary>
        /// COnverting the Password into securePassword
        /// </summary>
        /// <param name="password"></param>
        /// <returns></returns>
        public SecureString ConvertToSecureString(string password)
        {
            if (password == null)
                throw new ArgumentNullException("password");

            var securePassword = new SecureString();

            foreach (char c in password)
                securePassword.AppendChar(c);

            securePassword.MakeReadOnly();
            return securePassword;
        }
    }
}
