///////////////////////////////////////////////////////////////////////////////
// HTML Control and HTML Editor Sample
// Copyright 2003, Nikhil Kothari. All Rights Reserved.
//
// Provided as is, in sample form with no associated warranties.
// For more information on usage, see the accompanying
// License.txt file.
///////////////////////////////////////////////////////////////////////////////

namespace HtmlApp.Html {

    using System;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Drawing;

    /// <summary>
    /// Summary description for HtmlTextFormatting.
    /// </summary>
    public class HtmlTextFormatting {

        /// <summary>
        /// Array of the names of HTML formats that MSHTML accepts
        /// These need to be in the same order as the enum HtmlFormat
        /// </summary>
        private static readonly string[] formats =
            new string[] {
                             "Normal",               // Normal
                             "Formatted",            // PreFormatted
                             "Heading 1",            // Heading1
                             "Heading 2",            // Heading2
                             "Heading 3",            // Heading3
                             "Heading 4",            // Heading4
                             "Heading 5",            // Heading5
                             "Heading 6",            // Heading6
                             "Paragraph",            // Paragraph
                             "Numbered List",        // OrderedList
                             "Bulleted List"         // UnorderedList
                         };

        public HtmlEditor editor;

        /// <summary>
        /// Constructor which simply takes an HtmlEditor to interface with MSHTML
        /// </summary>
        /// <param name="target"></param>
        public HtmlTextFormatting(HtmlEditor editor) {
            this.editor = editor;
        }

        /// <summary>
        /// The background color of the current text
        /// </summary>
        public Color BackColor {
            get {
                //Query MSHTML, convert the result, and return the color
                return ConvertMSHTMLColor(editor.ExecResult(Interop.IDM_BACKCOLOR));
            }
            set {
                //Translate the color and execute the command
                string color = ColorTranslator.ToHtml(value);
                editor.Exec(Interop.IDM_BACKCOLOR, color);
            }
        }

        /// <summary>
        /// Indicates if the current text can be indented
        /// </summary>
        public bool CanIndent {
            get {
                return editor.IsCommandEnabled(Interop.IDM_INDENT);
            }
        }

        /// <summary>
        /// Indicates if the background color can be set
        /// </summary>
        public bool CanSetBackColor {
            get {
                return editor.IsCommandEnabled(Interop.IDM_BACKCOLOR);
            }
        }

        /// <summary>
        /// Indicates if the font face can be set
        /// </summary>
        public bool CanSetFontName {
            get {
                return editor.IsCommandEnabled(Interop.IDM_FONTNAME);
            }
        }

        /// <summary>
        /// Indicates if the font size can get set
        /// </summary>
        public bool CanSetFontSize {
            get {
                return editor.IsCommandEnabled(Interop.IDM_FONTSIZE);
            }
        }

        /// <summary>
        /// Indicates if the foreground color can be set
        /// </summary>
        public bool CanSetForeColor {
            get {
                return editor.IsCommandEnabled(Interop.IDM_FORECOLOR);
            }
        }

        /// <summary>
        /// Indicates if the Html format (eg ordered lists, paragraph, heading) can be set
        /// </summary>
        public bool CanSetHtmlFormat {
            get {
                return editor.IsCommandEnabled(Interop.IDM_BLOCKFMT);
            }
        }

        /// <summary>
        /// Indicates if the current text can be unindented
        /// </summary>
        public bool CanUnindent {
            get {
                return editor.IsCommandEnabled(Interop.IDM_OUTDENT);
            }
        }

        /// <summary>
        /// Gets and sets the font name of the current text
        /// </summary>
        public string FontName {
            get {
                return (editor.ExecResult(Interop.IDM_FONTNAME) as string);
            }
            set {
                editor.Exec(Interop.IDM_FONTNAME, value);
            }
        }

        /// <summary>
        /// Gets and sets the font size of the current text
        /// </summary>
        public HtmlFontSize FontSize {
            get {
                object o = editor.ExecResult(Interop.IDM_FONTSIZE);
                if (o == null) {
                    return HtmlFontSize.Medium;
                }
                else {
                    return (HtmlFontSize)o;
                }
            }
            set {
                editor.Exec(Interop.IDM_FONTSIZE, (int)value);
            }
        }

        /// <summary>
        /// The foreground color of the current text
        /// </summary>
        public Color ForeColor {
            get {
                //Query MSHTML, convert the result, and return the color
                return ConvertMSHTMLColor(editor.ExecResult(Interop.IDM_FORECOLOR));
            }
            set {
                //Translate the color and execute the command
                string color = ColorTranslator.ToHtml(value);
                editor.Exec(Interop.IDM_FORECOLOR, color);
            }
        }
        
        /// <summary>
        /// Converts an MSHTML color to a Frameworks color
        /// </summary>
        /// <param name="colorValue">object colorValue - The color value returned from MSHTML</param>
        /// <returns></returns>
        private Color ConvertMSHTMLColor(object colorValue) {
            if (colorValue != null) {
                Type colorType = colorValue.GetType();
                if (colorType == typeof(int)) {
                    //If the colorValue is an int, it's a Win32 color
                    return ColorTranslator.FromWin32((int)colorValue);
                }
                else if (colorType == typeof(string)) {
                    //Otherwise, it's a string, so convert that
                    return ColorTranslator.FromHtml((string)colorValue);
                }
                Debug.Fail("Unexpected color type : " + colorType.FullName);
            }
            return Color.Empty;
        }

        /// <summary>
        /// Gets the current state of the bold command (enabled and/or checked)
        /// </summary>
        /// <returns></returns>
        public HtmlCommandInfo GetBoldInfo() {
            return editor.GetCommandInfo(Interop.IDM_BOLD);
        }

        public HtmlFormat GetHtmlFormat() {
            string formatString = editor.ExecResult(Interop.IDM_BLOCKFMT) as string;
            if (formatString != null) {
                for (int i = 0; i < formats.Length; i++) {
                    if (formatString.Equals(formats[i])) {
                        return (HtmlFormat)i;
                    }
                }
            }
            return HtmlFormat.Normal;
        }

        /// <summary>
        /// Gets the current state of the italics command (enabled and/or checked)
        /// </summary>
        /// <returns></returns>
        public HtmlCommandInfo GetItalicsInfo() {
            return editor.GetCommandInfo(Interop.IDM_ITALIC);
        }

        /// <summary>
        /// Gets the current state of the strikethrough command (enabled and/or checked)
        /// </summary>
        /// <returns></returns>
        public HtmlCommandInfo GetStrikethroughInfo() {
            return editor.GetCommandInfo(Interop.IDM_STRIKETHROUGH);
        }

        /// <summary>
        /// Gets the current state of the Subscript command (enabled and/or checked)
        /// </summary>
        /// <returns></returns>
        public HtmlCommandInfo GetSubscriptInfo() {
            return editor.GetCommandInfo(Interop.IDM_SUBSCRIPT);
        }

        /// <summary>
        /// Gets the current state of the Superscript command (enabled and/or checked)
        /// </summary>
        /// <returns></returns>
        public HtmlCommandInfo GetSuperscriptInfo() {
            return editor.GetCommandInfo(Interop.IDM_SUPERSCRIPT);
        }

        /// <summary>
        /// Gets the current state of the Underline command (enabled and/or checked)
        /// </summary>
        /// <returns></returns>
        public HtmlCommandInfo GetUnderlineInfo() {
            return editor.GetCommandInfo(Interop.IDM_UNDERLINE);
        }

        /// <summary>
        /// Indents the current text
        /// </summary>
        public void Indent() {
            editor.Exec(Interop.IDM_INDENT);
        }
        /// <summary>
        /// Sets the HTML format (eg ordered list, paragraph, etc.) of the current text
        /// </summary>
        /// <param name="format"></param>
        public void SetHtmlFormat(HtmlFormat format) {
            editor.Exec(Interop.IDM_BLOCKFMT, formats[(int)format]);
        }

        /// <summary>
        /// Gets the current state of the bold command
        /// </summary>
        /// <returns></returns>
        public void ToggleBold() {
            editor.Exec(Interop.IDM_BOLD);
        }

        /// <summary>
        /// Toggles the current state of the italics command
        /// </summary>
        /// <returns></returns>
        public void ToggleItalics() {
            editor.Exec(Interop.IDM_ITALIC);
        }

        /// <summary>
        /// Toggles the current state of the strikethrough command
        /// </summary>
        /// <returns></returns>
        public void ToggleStrikethrough() {
            editor.Exec(Interop.IDM_STRIKETHROUGH);
        }

        /// <summary>
        /// Toggles the current state of the Subscript command
        /// </summary>
        /// <returns></returns>
        public void ToggleSubscript() {
            editor.Exec(Interop.IDM_SUBSCRIPT);
        }

        /// <summary>
        /// Toggles the current state of the Superscript command
        /// </summary>
        /// <returns></returns>
        public void ToggleSuperscript() {
            editor.Exec(Interop.IDM_SUPERSCRIPT);
        }

        /// <summary>
        /// Toggles the current state of the Underline command
        /// </summary>
        /// <returns></returns>
        public void ToggleUnderline() {
            editor.Exec(Interop.IDM_UNDERLINE);
        }

        /// <summary>
        /// Unindents the current text
        /// </summary>
        public void Unindent() {
            editor.Exec(Interop.IDM_OUTDENT);
        }
    }
}
