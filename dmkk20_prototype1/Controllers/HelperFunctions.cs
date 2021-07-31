using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace dmkk20_prototype1.Controllers
{
    public static class HelperFunctions
    {
        public static bool AssertNotEmptyText(string oldText, string newText)
        {
            return !(string.IsNullOrEmpty(oldText) || string.IsNullOrEmpty(newText));
        }
    }
}
