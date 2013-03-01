using System;
using System.Collections.Generic;
using System.Text;

namespace TimHeuer.PreviewHandlers
{
    public class CppFormat : Manoli.Utils.CSharpFormat.CSharpFormat
    {
        protected override string Keywords
        {
            get 
            {
                string str = base.Keywords + "auto static_cast reinterpret_cast dynamic_cast";
                return str;
            }
        }
    }
}
