// timheuer.com
// adapted from the MSDN Magazine samples January 2007 VOL 22 NO 1 edition

using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using TimHeuer.ManagedPreviewHandler;
using TimHeuer.PreviewHandlers.Properties;

namespace TimHeuer.PreviewHandlers
{
    [PreviewHandler("Source Code Preview Handler", ".cs;.vb;.sql;.js;.xaml;.xml;.htm;.html;.cpp;.h;.targets;.target", "{93E38957-78C4-40e2-9B1D-E202B43C6D23}")]
    [ProgId("TimHeuer.PreviewHandlers.CodePreviewHandler")]
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
                string formattedCode = FormatCode(previewCode, file.Extension.ToLowerInvariant());

                HtmlApp.Html.HtmlControl html = new HtmlApp.Html.HtmlControl();
                html.LoadHtml(string.Format("<html><head><style type=\"text/css\">{0}</style><body>{1}</body></html>", Resources.CssString, formattedCode));
                html.Dock = DockStyle.Fill;

                Controls.Add(html);
            }

            private string FormatCode(string sourceCode, string codeType)
            {
                string formatted = string.Empty;

                switch (codeType)
                {
                    case ".h":
                    case ".cpp":
                        CppFormat cpp = new CppFormat();
                        formatted = cpp.FormatCode(sourceCode);
                        break;
                    case ".cs":
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
                    case ".target":
                    case ".targets":
                        Manoli.Utils.CSharpFormat.HtmlFormat xml = new Manoli.Utils.CSharpFormat.HtmlFormat();
                        formatted = xml.FormatCode(sourceCode);
                        break;
                }

                return formatted;
            }
        }
    }
}
