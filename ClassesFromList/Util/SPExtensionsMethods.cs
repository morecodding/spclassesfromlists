using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace ClassesFromList.Util
{
    static class SPExtensionsMethods
    {
        /// <summary>
        /// Converte uma string comum para uma string do tipo SecureString
        /// </summary>
        /// <param name="pass">string comum</param>
        /// <returns>String Segura</returns>
        public static SecureString ToSecureString(this string pass)
        {
            var secure = new SecureString();
            if (pass.Length > 0)
                foreach (var c in pass.ToCharArray()) secure.AppendChar(c);
            return secure;
        }
    }
}
