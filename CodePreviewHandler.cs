// timheuer.com
// adapted from the MSDN Magazine samples January 2007 VOL 22 NO 1 edition

using System;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Globalization;
using MsdnMag;

namespace SmilingGoat.PreviewHandlers
{
    [PreviewHandler("Source Code Preview Handler", ".cs;.vb;.sql;.js", "{93E38957-78C4-40e2-9B1D-E202B43C6D23}")]
    [ProgId("SmilingGoat.PreviewHandlers.CodePreviewHandler")]
    [Guid("0E1B4233-AEB5-4c5b-BF31-21766492B301")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    public sealed class CodePreviewHandler : FileBasedPreviewHandler
    {
        protected override PreviewHandlerControl CreatePreviewHandlerControl()
        {
            return new CodePreviewHandlerControl();
        }

        private sealed class CodePreviewHandlerControl : FileBasedPreviewHandlerControl
        {
            public override void Load(FileInfo file)
            {
                StreamReader rdr = file.OpenText();
                string previewCode = rdr.ReadToEnd();
                string formattedCode = FormatCode(previewCode, file.Extension);

                HtmlApp.Html.HtmlControl html = new HtmlApp.Html.HtmlControl();
                html.LoadHtml(string.Format("<html><head><style type=\"text/css\">{0}</style><body>{1}</body></html>", SmilingGoat.PreviewHandlers.Properties.Resources.CssString, formattedCode));
                html.Dock = DockStyle.Fill;

                Controls.Add(html);
            }

            internal string FormatCode(string sourceCode, string codeType)
            {
                string formatted = string.Empty;

                switch (codeType)
                {
                    case ".cs":
                    case ".cpp":
                    case ".h":
                        Manoli.Utils.CSharpFormat.CSharpFormat cs = new Manoli.Utils.CSharpFormat.CSharpFormat();
                        formatted = cs.FormatCode(sourceCode);
                        break;
                    case ".vb":
                        Manoli.Utils.CSharpFormat.VisualBasicFormat vb = new Manoli.Utils.CSharpFormat.VisualBasicFormat();
                        formatted = vb.FormatCode(sourceCode);
                        break;
                    case ".js":
                        Manoli.Utils.CSharpFormat.JavaScriptFormat js = new Manoli.Utils.CSharpFormat.JavaScriptFormat();
                        formatted = js.FormatCode(sourceCode);
                        break;
                    case ".sql":
                        Manoli.Utils.CSharpFormat.TsqlFormat sql = new Manoli.Utils.CSharpFormat.TsqlFormat();
                        formatted = sql.FormatCode(sourceCode);
                        break;
                    case ".xaml":
                    case ".xml":
                    case ".html":
                    case ".htm":
                        Manoli.Utils.CSharpFormat.TsqlFormat xml = new Manoli.Utils.CSharpFormat.HtmlFormat();
                        formatted = xml.FormatCode(sourceCode);
                        break;
                    default:
                        break;
                }

                return formatted;
            }
        }
    }
}
