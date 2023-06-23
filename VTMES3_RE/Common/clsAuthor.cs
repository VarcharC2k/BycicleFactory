using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VTMES3_RE.Common
{
    public class clsAuthor
    {
        public string MenuID = "";
        public string ParentMenuID = "";
        public string AuthorCode = "";

        public clsAuthor(string _menuID, string _parentMenuID, string _authorCode)
        {
            MenuID = _menuID;
            ParentMenuID = _parentMenuID;
            AuthorCode = _authorCode;
        }
    }
}
