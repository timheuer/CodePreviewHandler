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

    /// <summary>
    /// Summary description for HtmlDocument.
    /// </summary>
    public class HtmlDocument {

        private HtmlEditor _editor;

        public HtmlDocument(HtmlEditor editor) {
            _editor = editor;
        }

        /// <summary>
        /// Indicates if a button can be inserted
        /// </summary>
        public bool CanInsertButton {
            get {
                return _editor.IsCommandEnabled(Interop.IDM_BUTTON);
            }
        }

        /// <summary>
        /// Indicates if a listbox can be inserted
        /// </summary>
        public bool CanInsertListBox {
            get {
                return _editor.IsCommandEnabled(Interop.IDM_LISTBOX);
            }
        }

        /// <summary>
        /// Indicates if HTML can be inserted
        /// </summary>
        public bool CanInsertHtml {
            get {
                if (Selection.Type == HtmlSelectionType.ElementSelection) {
                    //If this is a control range, we can only insert HTML if we're in a div or span
                    Interop.IHtmlControlRange controlRange = (Interop.IHtmlControlRange)Selection.MSHTMLSelection;
                    int selectedItemCount = controlRange.GetLength();
                    if (selectedItemCount == 1) {
                        Interop.IHTMLElement element = controlRange.Item(0);
                        if ((String.Compare(element.GetTagName(), "div", true) == 0) ||
                            (String.Compare(element.GetTagName(), "td", true) == 0)) {
                            return true;
                        }
                    }
                }
                else {
                    //If this is a text range, we can definitely insert HTML
                    return true;
                }
                return false;
            }
        }

        /// <summary>
        /// Indicates if an hyperlink can be inserted
        /// </summary>
        public bool CanInsertHyperlink {
            get {
                if (((Selection.Type == HtmlSelectionType.TextSelection) || (Selection.Type == HtmlSelectionType.Empty)) && 
                    (Selection.Length == 0)) {
                    return CanInsertHtml;
                }
                else {
                    return _editor.IsCommandEnabled(Interop.IDM_HYPERLINK);
                }
            }
        }

        /// <summary>
        /// Indicates if a radio button can be inserted
        /// </summary>
        public bool CanInsertRadioButton {
            get {

                return _editor.IsCommandEnabled(Interop.IDM_RADIOBUTTON);
            }
        }

        /// <summary>
        /// Indicates if a text area can be inserted
        /// </summary>
        public bool CanInsertTextArea {
            get {
                return _editor.IsCommandEnabled(Interop.IDM_TEXTAREA);
            }
        }

        /// <summary>
        /// Indicates if a textbox can be inserted
        /// </summary>
        public bool CanInsertTextBox {
            get {
                return _editor.IsCommandEnabled(Interop.IDM_TEXTBOX);
            }
        }

        /// <summary>
        /// The current selection in the editor
        /// </summary>
        protected HtmlSelection Selection {
            get {
                return _editor.Selection;
            }
        }

        /// <summary>
        /// Inserts a button
        /// </summary>
        public void InsertButton() {
            _editor.Exec(Interop.IDM_BUTTON);
        }

        /// <summary>
        /// Inserts the specified string into the html over the current selection
        /// </summary>
        /// <param name="html"></param>
        public void InsertHtml(string html) {
            Selection.SynchronizeSelection();
            if (Selection.Type == HtmlSelectionType.ElementSelection) {
                //If it's a control range, we can only insert if we are in a div or td
                Interop.IHtmlControlRange controlRange = (Interop.IHtmlControlRange)Selection.MSHTMLSelection;
                int selectedItemCount = controlRange.GetLength();
                if (selectedItemCount == 1) {
                    Interop.IHTMLElement element = controlRange.Item(0);
                    if ((String.Compare(element.GetTagName(), "div", true) == 0) ||
                        (String.Compare(element.GetTagName(), "td", true) == 0)) {
                        element.InsertAdjacentHTML("beforeEnd", html);
                    }
                }
            }
            else {
                Interop.IHTMLTxtRange textRange = (Interop.IHTMLTxtRange)Selection.MSHTMLSelection;
                textRange.PasteHTML(html);
            }
        }

        /// <summary>
        /// Inserts a hyperlink with the specified URL and description
        /// </summary>
        /// <param name="url"></param>
        /// <param name="description"></param>
        public void InsertHyperlink(string url, string description) {
            Selection.SynchronizeSelection();
            if (url == null) {
                try {
                    _editor.ExecPrompt(Interop.IDM_HYPERLINK);
                }
                catch {} // don't care if it fails!
            }
            else {
                if (((Selection.Type == HtmlSelectionType.TextSelection) || (Selection.Type == HtmlSelectionType.Empty)) && 
                    (Selection.Length == 0)) {
                    InsertHtml("<a href=\""+url+"\">"+description+"</a>");
                    /*Interop.IHTMLTxtRange textRange = (Interop.IHTMLTxtRange)Selection.MSHTMLSelection;
                    textRange.PasteHTML("<a href=\""+url+"\">"+description+"</a>");*/
                }
                else {
                    _editor.Exec(Interop.IDM_HYPERLINK, url);
                }
            }
        }

        /// <summary>
        /// Inserts a list box
        /// </summary>
        public void InsertListBox() {
            _editor.Exec(Interop.IDM_LISTBOX);
        }

        /// <summary>
        /// Inserts a radio button
        /// </summary>
        public void InsertRadioButton() {
            _editor.Exec(Interop.IDM_RADIOBUTTON);
        }

        /// <summary>
        /// Inserts a text area
        /// </summary>
        public void InsertTextArea() {
            _editor.Exec(Interop.IDM_TEXTAREA);
        }

        /// <summary>
        /// Inserts a text box
        /// </summary>
        public void InsertTextBox() {
            _editor.Exec(Interop.IDM_TEXTBOX);
        }
    }
}
