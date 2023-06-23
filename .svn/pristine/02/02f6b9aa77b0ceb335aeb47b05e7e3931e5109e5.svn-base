using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace VTMES3_RE.Common
{
    public class clsCommon
    {
        public static string getString(object obj)
        {
            if (Convert.IsDBNull(obj))
            {
                return "";
            }
            else
            {
                if (obj.GetType().Name == "DateTime")
                {
                    DateTime dtime = (DateTime)obj;
                    return dtime.ToShortDateString();
                }
                else
                {
                    return obj.ToString();
                }

            }
        }

        public static bool IsLetter(char c)
        {
            return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
        }

        public static bool IsDigit(char c)
        {
            return c >= '0' && c <= '9';
        }

        public static bool IsSymbol(char c)
        {
            return c > 32 && c < 127 && !IsDigit(c) && !IsLetter(c);
        }

        public static bool IsValidPassword(string password)
        {
            return
               password.Any(c => IsLetter(c)) &&
               password.Any(c => IsDigit(c)) &&
               password.Any(c => IsSymbol(c));
        }

        public static bool ValidatePassword(string password)
        {
            const int MIN_LENGTH = 8;
            const int MAX_LENGTH = 20;

            if (password == null) throw new ArgumentNullException();

            bool meetsLengthRequirements = password.Length >= MIN_LENGTH && password.Length <= MAX_LENGTH;
            bool isValid = false;
            int hasUpperCaseLetter = 0;
            int hasLowerCaseLetter = 0;
            int hasDecimalDigit = 0;
            int hasSymbolChar = 0;

            if (meetsLengthRequirements)
            {
                foreach (char c in password)
                {
                    if (char.IsUpper(c)) hasUpperCaseLetter = 1;
                    else if (char.IsLower(c)) hasLowerCaseLetter = 1;
                    else if (char.IsDigit(c)) hasDecimalDigit = 1;
                    else if (IsSymbol(c)) hasSymbolChar = 1;
                }
            }

            int validCount = hasUpperCaseLetter + hasLowerCaseLetter + hasDecimalDigit + hasSymbolChar;

            if (validCount >= 3 && password.Length >= 8)
            {
                isValid = true;
            }
            else if (validCount >= 2 && password.Length >= 10)
            {
                isValid = true;
            }

            return isValid;

        }

        // SHA256  256bit 암호화
        public static string SHA256Hash(string Data)
        {
            SHA256 sha = new SHA256Managed();
            byte[] hash = sha.ComputeHash(Encoding.ASCII.GetBytes(Data));

            StringBuilder stringBuilder = new StringBuilder();

            foreach (byte b in hash)
            {
                stringBuilder.AppendFormat("{0:x2}", b);
            }
            return stringBuilder.ToString();
        }
    }
}
