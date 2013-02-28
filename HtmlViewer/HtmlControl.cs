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
    using System.Collections;
    using System.Collections.Specialized;
    using System.Diagnostics;
    using System.Drawing;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Windows;
    using System.Windows.Forms;

    using STATSTG = Interop.STATSTG;

    /// <summary>
    /// An HTML rendering control based on MSHTML
    /// </summary>
    public class HtmlControl : Control {

        private static readonly object EventShowContextMenu = new object();
        private static readonly object _readyStateCompleteEvent = new object();

        private bool _scrollBarsEnabled;
        private bool _flatScrollBars;
        private bool _border3d;
        private bool _scriptEnabled;
        private bool _allowInPlaceNavigation;
        private bool _fullDocumentMode;

        private bool _firstActivation;
        private bool _isReady;
        private bool _isCreated;

        //These allow a user to load the document before displaying
        private bool _loadDesired;
        private string _desiredContent;
        private string _desiredUrl;

        private bool _focusDesired;

        private string _url;
        private object _scriptObject;

        private MSHTMLSite _site;

        private static IDictionary _urlMap;

        /// <summary>
        /// </summary>
        public HtmlControl() : this(true) {
        }

        /// <summary>
        /// </summary>
        /// <param name="fullDocumentMode"></param>
        public HtmlControl(bool fullDocumentMode) {
            _firstActivation = true;
            _fullDocumentMode = fullDocumentMode;

            // Scroll bars should be enabled by default
            _scrollBarsEnabled = true;
        }

        public bool AllowInPlaceNavigation {
            get {
                return _allowInPlaceNavigation;
            }
            set {
                _allowInPlaceNavigation = value;
            }
        }

        /// <summary>
        /// Indicates if the Interop.HTMLDocument2 is created
        /// </summary>
        protected bool IsCreated {
            get {
                return _isCreated;
            }
        }

        internal bool IsFullDocumentMode {
            get {
                return _fullDocumentMode;
            }
        }

        /// <summary>
        /// Indicates if the control is ready for use
        /// </summary>
        public bool IsReady {
            get {
                return _isReady;
            }
        }

        protected internal Interop.IHTMLDocument2 MSHTMLDocument {
            get {
                return _site.MSHTMLDocument;
            }
        }

        protected internal Interop.IOleCommandTarget CommandTarget {
            get {
                return _site.MSHTMLCommandTarget;
            }
        }

        public bool Border3d {
            get {
                return _border3d;
            }
            set {
                _border3d = value;
            }
        }

        /// <summary>
        /// Indicates if the current selection can be copied
        /// </summary>
        public bool CanCopy {
            get {
                return IsCommandEnabled(Interop.IDM_COPY);
            }
        }

        /// <summary>
        /// Indicates if the current selection can be cut
        /// </summary>
        public bool CanCut {
            get {
                return IsCommandEnabled(Interop.IDM_CUT);
            }
        }

        /// <summary>
        /// Indicates if the current selection can be pasted to
        /// </summary>
        public bool CanPaste {
            get {
                return IsCommandEnabled(Interop.IDM_PASTE);
            }
        }

        /// <summary>
        /// Indicates if the editor can redo
        /// </summary>
        public bool CanRedo {
            get {
                return IsCommandEnabled(Interop.IDM_REDO);
            }
        }

        /// <summary>
        /// Indicates if the editor can undo
        /// </summary>
        public bool CanUndo {
            get {
                return IsCommandEnabled(Interop.IDM_UNDO);
            }
        }

        public bool FlatScrollBars {
            get {
                return _flatScrollBars;
            }
            set {
                _flatScrollBars = value;
            }
        }


        public event EventHandler ReadyStateComplete {
            add {
                Events.AddHandler(_readyStateCompleteEvent, value);
            }
            remove {
                Events.RemoveHandler(_readyStateCompleteEvent, value);
            }
        }

        public bool ScriptEnabled {
            get {
                return _scriptEnabled;
            }
            set {
                _scriptEnabled = value;
            }
        }

        public object ScriptObject {
            get {
                return _scriptObject;
            }
            set {
                _scriptObject = value;
            }
        }

        public bool ScrollBarsEnabled {
            get {
                return _scrollBarsEnabled;
            }
            set {
                _scrollBarsEnabled = value;
            }
        }

        /// <summary>
        /// Gets the url of the document contained in the control
        /// </summary>
        public virtual string Url {
            get {
                return _url;
            }
        }

        internal static IDictionary UrlMap {
            get {
                if (_urlMap == null) {
                    _urlMap = new HybridDictionary(true);
                }
                return _urlMap;
            }
        }

        /// <summary>
        /// Copy the current selection
        /// </summary>
        public void Copy() {
            if (!CanCopy) {
                throw new Exception("HtmlControl.Copy : Not in able to copy the current selection!");
            }
            Exec(Interop.IDM_COPY);
        }

        protected virtual string CreateHtmlContent(string content, string style) {
            return "<html><head>" + style + "</head><body>" + content + "</body></html>";
        }

        /// <summary>
        /// Cut the current selection
        /// </summary>
        public void Cut() {
            if (!CanCut) {
                throw new Exception("HtmlControl.Cut : Not in able to cut the current selection!");
            }
            Exec(Interop.IDM_CUT);
        }

        /// <internalonly/>
        protected override void Dispose(bool disposing) {
            if (disposing) {
                if (_url != null) {
                    UrlMap[_url] = null;
                }
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// Executes the specified command in MSHTML
        /// </summary>
        /// <param name="command"></param>
        protected internal void Exec(int command) {
            Exec(command, null);
        }

        /// <summary>
        /// Executes the specified command in MSHTML with the specified argument
        /// </summary>
        /// <param name="command"></param>
        protected internal void Exec(int command, object argument) {
            object[] args = new object[] { argument };

            //Execute the command
            int hr = CommandTarget.Exec(ref Interop.Guid_MSHTML, command, Interop.OLECMDEXECOPT.OLECMDEXECOPT_DONTPROMPTUSER, args, null);
            if (hr != Interop.S_OK) {
                throw new Exception("MSHTMLHelper.Exec : Command "+command+" did not return S_OK");
            }
        }

        /// <summary>
        /// Executes the specified command in MSHTML with the specified argument
        /// </summary>
        /// <param name="command"></param>
        protected internal void ExecPrompt(int command) {
            object[] args = new object[] { null };

            //Execute the command
            int hr = CommandTarget.Exec(ref Interop.Guid_MSHTML, command, Interop.OLECMDEXECOPT.OLECMDEXECOPT_PROMPTUSER, args, null);
            if (hr != Interop.S_OK) {
                throw new Exception("MSHTMLHelper.Exec : Command "+command+" did not return S_OK");
            }
        }

        /// <summary>
        /// Executes the specified command in MSHTML and returns the result
        /// </summary>
        /// <param name="command"></param>
        /// <returns>object result - The result of the command</returns>
        protected internal object ExecResult(int command) {
            object[] retVal = new object[1];

            //Execute the command
            int hr = CommandTarget.Exec(ref Interop.Guid_MSHTML, command, Interop.OLECMDEXECOPT.OLECMDEXECOPT_DONTPROMPTUSER, null, retVal);
            if (hr != Interop.S_OK) {
                throw new Exception("MSHTMLHelper.ExecResult : Command "+command+" did not return S_OK");
            }
            return retVal[0];
        }

        /// <summary>
        /// Queries MSHTML for the command info (enabled and checked) for the specified command
        /// </summary>
        /// <param name="command"></param>
        /// <returns></returns>
        protected internal HtmlCommandInfo GetCommandInfo(int command) {
            //First query the command target for the command status
            int info;
            Interop.tagOLECMD oleCommand = new Interop.tagOLECMD();

            //Create an tagOLDCMD to store the command and receive the result
            oleCommand.cmdID = command;
            int hr = CommandTarget.QueryStatus(ref Interop.Guid_MSHTML, 1, oleCommand, 0);
            Debug.Assert(hr == Interop.S_OK,"IOleCommand.QueryStatus did not return S_OK");

            //Then translate the response from the command status
            //We can just right shift by one to eliminate the supported flag from OLECMDF
            info = oleCommand.cmdf;
            //REVIEW: Do we want to do a mapping instead of playing with the bits?
            return (HtmlCommandInfo)(info>>1) & (HtmlCommandInfo.Enabled | HtmlCommandInfo.Checked);
        }

        /// <summary>
        /// Indicates if the specified command is checked
        /// </summary>
        /// <param name="command"></param>
        /// <returns></returns>
        protected internal bool IsCommandChecked(int command) {
            return ((GetCommandInfo(command) & HtmlCommandInfo.Checked) != 0);
        }

        /// <summary>
        /// Indicates if the specified command is enabled
        /// </summary>
        /// <param name="command"></param>
        /// <returns></returns>
        protected internal bool IsCommandEnabled(int command) {
            return ((GetCommandInfo(command) & HtmlCommandInfo.Enabled) != 0);
        }

        /// <summary>
        /// Indicates if the specified command is enabled
        /// </summary>
        /// <param name="command"></param>
        /// <returns></returns>
        protected internal bool IsCommandEnabledAndChecked(int command) {
            HtmlCommandInfo info = GetCommandInfo(command);
            return (((info & HtmlCommandInfo.Enabled) != 0) && ((info & HtmlCommandInfo.Checked) != 0));
        }

        /// <summary>
        /// Loads HTML content from a stream into this control
        /// </summary>
        /// <param name="stream"></param>
        public void LoadHtml(Stream stream) {
            if (stream == null) {
                throw new ArgumentNullException("LoadHtml : You must specify a non-null stream for content");
            }
            StreamReader reader = new StreamReader(stream);
            LoadHtml(reader.ReadToEnd());
        }

        /// <summary>
        /// Loads HTML content from a string into this control
        /// </summary>
        /// <param name="content"></param>
        public void LoadHtml(string content) {
            LoadHtml(content, null, null);
        }

        public void LoadHtml(string content, string url) {
            LoadHtml(content, url, null);
        }

        //REVIEW: Add a load method for stream and url

        /// <summary>
        /// Loads HTML content from a string into this control identified by the specified URL.
        /// If MSHTML has not yet been created, the loading is postponed until MSHTML has been created.
        /// </summary>
        /// <param name="content"></param>
        /// <param name="url"></param>
        public void LoadHtml(string content, string url, string style) {
            if (content == null) {
                content = "";
            }

            if (!_isCreated) {
                _desiredContent = content;
                _desiredUrl = url;
                _loadDesired = true;
                return;
            }

            if (_fullDocumentMode == false) {
                content = CreateHtmlContent(content, style);
            }

            OnBeforeLoad();

            Interop.IStream stream = null;

            //First we create a COM stream
            IntPtr hglobal = Marshal.StringToHGlobalUni(content);
            Interop.CreateStreamOnHGlobal(hglobal, true, out stream);

            // Initialize a new document if there is nothing to load
            if (stream == null) {
                Interop.IPersistStreamInit psi = (Interop.IPersistStreamInit)_site.MSHTMLDocument;
                Debug.Assert(psi != null, "Expected IPersistStreamInit");
                psi.InitNew();
                psi = null;
            }
            else {
                Interop.IHTMLDocument2 document = _site.MSHTMLDocument;
                //If there is no specified URL load the document from the stream
                if (url == null) {
                    Interop.IPersistStreamInit psi = (Interop.IPersistStreamInit)document;
                    Debug.Assert(psi != null, "Expected IPersistStreamInit");
                    psi.Load(stream);
                    psi = null;
                }
                else {
                    //Otherwise we create a moniker and load the stream to that moniker
                }
            }
            _url = url;

            OnAfterLoad();
        }

        /// <summary>
        /// Allow editors to perform actions after HTML content is loaded to the control
        /// </summary>
        protected virtual void OnAfterLoad() {
        }

        /// <summary>
        /// </summary>
        protected virtual void OnAfterSave() {
        }

        /// <summary>
        /// Allow editors to perform actions before HTML content is loaded to the control
        /// </summary>
        protected virtual void OnBeforeLoad() {
        }

        /// <summary>
        /// </summary>
        protected virtual void OnBeforeSave() {
        }

        /// <summary>
        /// Allow editors to perform actions when the MSHTML document is created
        /// and before it's activated
        /// </summary>
        /// <param name="args"></param>
        protected virtual void OnCreated(EventArgs args) {
        }

        /// <summary>
        /// On focus, we have to also return focus to MSHTML
        /// </summary>
        protected override void OnGotFocus(EventArgs e) {
            base.OnGotFocus(e);

            //TODO: Fill this in with the right code
            if (IsReady) {
                // REVIEW: Not sure if we need to do this...  It seems to work without it.
                //_site.ActivateMSHTML();
                _site.SetFocus();
            }
            else {
                _focusDesired = true;
            }
        }

        /// <summary>
        /// We can only activate the MSHTML after our handle has been created,
        /// so upon creating the handle, we create and activate Interop.
        ///
        /// If LoadHtml was called prior to this, we do the loading now
        /// </summary>
        /// <param name="args"></param>
        protected override void OnHandleCreated(EventArgs args) {
            base.OnHandleCreated(args);
            if (_firstActivation) {
                _site = new MSHTMLSite(this);
                _site.CreateMSHTML();

                _isCreated = true;

                OnCreated(new EventArgs());

                _site.ActivateMSHTML();
                _firstActivation = false;

                if (_loadDesired) {
                    LoadHtml(_desiredContent, _desiredUrl);
                    _loadDesired = false;
                }
            }
        }

        /// <summary>
        /// Called when the control has just become ready
        /// </summary>
        /// <param name="e"></param>
        protected internal virtual void OnReadyStateComplete(EventArgs e) {
            _isReady = true;

            EventHandler handler = (EventHandler)Events[_readyStateCompleteEvent];
            if (handler != null) {
                handler(this, e);
            }

            if (_focusDesired) {
                _focusDesired = false;
                _site.ActivateMSHTML();
                _site.SetFocus();
            }
        }

        /// <summary>
        /// Cut the current selection
        /// </summary>
        public void Paste() {
            if (!CanPaste) {
                throw new Exception("HtmlControl.Paste : Not in able to paste the current selection!");
            }
            Exec(Interop.IDM_PASTE);
        }

        /// <summary>
        /// We need to process keystrokes in order to pass them to MSHTML
        /// </summary>
        public override bool PreProcessMessage(ref Message m) {
            bool handled = false;
            if ((m.Msg >= Interop.WM_KEYFIRST) && (m.Msg <= Interop.WM_KEYLAST)) {
                // If it's a key down, first see if the key combo is a command key

                if (m.Msg == Interop.WM_KEYDOWN) {
                    handled = ProcessCmdKey(ref m, (Keys)(int)m.WParam | ModifierKeys);
                }

                if (!handled) {
                    int keyCode = (int)m.WParam;
                    // Don't let Trident eat Ctrl-PgUp/PgDn
                    if (((keyCode != (int)Keys.PageUp) && (keyCode != (int)Keys.PageDown)) || ((ModifierKeys & Keys.Control) == 0)) {
                        Interop.COMMSG cm = new Interop.COMMSG();

                        cm.hwnd = m.HWnd;
                        cm.message = m.Msg;
                        cm.wParam = m.WParam;
                        cm.lParam = m.LParam;
                        handled = _site.TranslateAccelarator(cm);
                    }
                    else {
                        // WndProc for Ctrl-PgUp/PgDn is never called so call it directly here
                        WndProc(ref m);
                        handled = true;
                    }
                }
            }

            if (!handled) {
                handled = base.PreProcessMessage(ref m);
            }

            return handled;
        }

        public void Redo() {
            if (!CanRedo) {
                throw new Exception("HtmlControl.Redo : Not in able to redo!");
            }
            Exec(Interop.IDM_REDO);
        }

        /// <summary>
        /// Saves the HTML contained in control to a string and return it.
        /// </summary>
        /// <returns>string - The HTML in the control</returns>
        public string SaveHtml() {
            if (!IsCreated) {
                throw new Exception("HtmlControl.SaveHtml : No HTML to save!");
            }

            string content = String.Empty;

            try {
                OnBeforeSave();

                Interop.IHTMLDocument2 document = _site.MSHTMLDocument;

                if (_fullDocumentMode) {
                    // First save the document to a stream
                    Interop.IPersistStreamInit psi = (Interop.IPersistStreamInit)document;
                    Debug.Assert(psi != null, "Expected IPersistStreamInit");

                    Interop.IStream stream = null;
                    Interop.CreateStreamOnHGlobal(Interop.NullIntPtr, true, out stream);

                    psi.Save(stream, 1);

                    // Now copy the stream to the string
                    STATSTG stat = new STATSTG();
                    stream.Stat(stat, 1);
                    int length = (int)stat.cbSize;
                    byte[] bytes = new byte[length];

                    IntPtr hglobal;
                    Interop.GetHGlobalFromStream(stream, out hglobal);
                    Debug.Assert(hglobal != Interop.NullIntPtr, "Failed in GetHGlobalFromStream");

                    // First copy the stream to a byte array
                    IntPtr pointer = Interop.GlobalLock(hglobal);
                    if (pointer != Interop.NullIntPtr) {
                        Marshal.Copy(pointer, bytes, 0, length);

                        Interop.GlobalUnlock(hglobal);

                        // Then create the string from the byte array (use a StreamReader to eat the preamble in the UTF8 encoding case)
                        StreamReader streamReader = null;
                        try {
                            streamReader = new StreamReader(new MemoryStream(bytes), Encoding.Default);
                            content = streamReader.ReadToEnd();
                        }
                        finally {
                            if (streamReader != null) {
                                streamReader.Close();
                            }
                        }
                    }
                }
                else {
                    // Save only the contents of the <body> tag
                    Interop.IHTMLElement bodyElement = document.GetBody();
                    Debug.Assert(bodyElement != null, "Could not get BODY element from document");

                    if (bodyElement != null) {
                        content = SavePartialHtml(Element.GetWrapperFor(bodyElement, this));
                    }
                }
            }
            catch (Exception e) {
                Debug.Fail("HtmlControl.SaveHtml" + e.ToString());
                content = String.Empty;
            }
            finally {
                OnAfterSave();
            }

            if (content == null) {
                content = String.Empty;
            }
            return content;
        }

        // REVIEW: Come up with better names to unify _fullDocumentMode, CreateHtmlContent, and SavePartialHtml
        protected virtual string SavePartialHtml(Element bodyElement) {
            return bodyElement.InnerHtml;
        }

        /// <summary>
        /// Saves the HTML contained in the control to a stream
        /// </summary>
        /// <param name="stream"></param>
        public void SaveHtml(Stream stream) {
            if (stream == null) {
                throw new ArgumentNullException("SaveHtml : Must specify a non-null stream to which to save");
            }

            string content = SaveHtml();

            StreamWriter writer = new StreamWriter(stream, Encoding.UTF8);
            writer.Write(content);
            writer.Flush();
        }

        public void Undo() {
            if (!CanUndo) {
                throw new Exception("HtmlControl.Undo : Not in able to undo!");
            }
            Exec(Interop.IDM_UNDO);
        }
    }
}
