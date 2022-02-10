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
    using System.Diagnostics;
    using System.ComponentModel.Design;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    /// <summary>
    /// Summary description for HtmlEditor.
    /// </summary>
    public class HtmlEditor : HtmlControl {
        private bool _absolutePositioningEnabled;
        private bool _absolutePositioningDesired;

        private bool _bordersVisible;
        private bool _bordersDesired;

        private bool _designModeEnabled;
        private bool _designModeDesired;

        private bool _multipleSelectionEnabled;
        private bool _multipleSelectionDesired;

        private HtmlTextFormatting textFormatting;
        private HtmlDocument document;
        private HtmlSelection _selection;

        Interop.IPersistStreamInit _persistStream;

        private Interop.IOleUndoManager _undoManager;

        private Interop.IHTMLEditServices _editServices;

        public HtmlEditor() : this(true) {
        }

        internal HtmlEditor(bool fullDocumentMode) : base(fullDocumentMode) {
        }

        /// <summary>
        /// Enables or disables absolute position for the entire editor
        /// </summary>
        public bool AbsolutePositioningEnabled {
            get {
                return _absolutePositioningEnabled;
            }
            set {
                //If the control isn't ready to be put into abs pos mode,
                //set a flag and put it in abs pos mode when it is ready
                _absolutePositioningDesired = value;
                if (!IsCreated) {
                    return;
                }
                else {
                    //Turn abs pos mode on or off depending on the new value
                    _absolutePositioningEnabled = value;
                    object[] args = new object[] { _absolutePositioningEnabled };
                    Exec(Interop.IDM_2D_POSITION, args);
                }
            }
        }

        public bool BordersVisible {
            get {
                return _bordersVisible;
            }
            set {
                _bordersDesired = value;
                if (!IsReady) {
                    return;
                }
                if (_bordersVisible != _bordersDesired) {
                    _bordersVisible = value;
                    object[] args = new object[] { _bordersVisible };
                    Exec(Interop.IDM_SHOWZEROBORDERATDESIGNTIME,args);
                }
            }
        }

        /// <summary>
        /// Indicates if the editor is in design mode
        /// Also places MSHTML into design mode if set to true
        /// </summary>
        public bool DesignModeEnabled {
            get {
                return _designModeEnabled;
            }
            set {
                //Only execute this if we aren't already in design mode
                if (_designModeEnabled != value) {
                    //If the control isn't ready to be put into design mode,
                    //set a flag and put it in design mode when it is ready
                    if (!IsCreated) {
                        _designModeDesired = value;
                    }
                    else {
                        //Turn design mode on or off depending on the new value
                        _designModeEnabled = value;
                        MSHTMLDocument.SetDesignMode((_designModeEnabled ? "on" : "off"));
                    }
                }
            }
        }

        /// <summary>
        /// Returns the document object for doing insertions
        /// </summary>
        public HtmlDocument Document {
            get {
                if (!IsReady) {
                    throw new Exception("HtmlDocument not ready yet!");
                }
                if (document == null) {
                    document = new HtmlDocument(this);
                }
                return document;
            }
        }

        private Interop.IHTMLEditServices MSHTMLEditServices {
            get {
                if (_editServices==null) {
                    Interop.IServiceProvider serviceProvider = MSHTMLDocument as Interop.IServiceProvider;
                    Debug.Assert(serviceProvider != null);
                    Guid shtmlGuid = new Guid(0x3050f7f9,0x98b5,0x11cf,0xbb,0x82,0x00,0xaa,0x00,0xbd,0xce,0x0b);
                    Guid intGuid = (typeof(Interop.IHTMLEditServices)).GUID;

                    IntPtr editServicePtr = Interop.NullIntPtr;
                    int hr = serviceProvider.QueryService(ref shtmlGuid, ref intGuid, out editServicePtr);
                    Debug.Assert((hr == Interop.S_OK) && (editServicePtr != Interop.NullIntPtr), "Did not get IHTMLEditService");
                    if ((hr == Interop.S_OK) && (editServicePtr != Interop.NullIntPtr)) {
                        _editServices = (Interop.IHTMLEditServices) Marshal.GetObjectForIUnknown(editServicePtr);
                        Marshal.Release(editServicePtr);
                    }
                }
                return _editServices;
            }
        }
        /// <summary>
        /// Indicates if the contents of the editor have been modified
        /// </summary>
        public virtual bool IsDirty {
            get {
                if (DesignModeEnabled && IsReady) {
                    if (_persistStream != null) {
                        //TODO: After a load, this is no longer true, what do we do???
                        if (_persistStream.IsDirty() == Interop.S_OK) {
                            return true;
                        }
                    }
                }
                return false;
            }
        }

        /// <summary>
        /// Indicates if multiple selection is enabled in the editor
        /// Also places MSHTML into multiple selection mode if set to true
        /// </summary>
        public bool MultipleSelectionEnabled {
            get {
                return _multipleSelectionEnabled;
            }
            set {
                //If the control isn't ready yet, postpone setting multiple selection
                _multipleSelectionDesired = value;
                if (!IsReady) {
                    return;
                }
                else {
                    //Create an objects array to pass parameters to the MSHTML command target
                    _multipleSelectionEnabled = value;
                    object[] args = new object[] { _multipleSelectionEnabled };
                    int hr = CommandTarget.Exec(ref Interop.Guid_MSHTML, Interop.IDM_MULTIPLESELECTION, 0, args, null);
                    Debug.Assert(hr == Interop.S_OK);
                }
            }
        }

        /// <summary>
        /// The current selection in the editor
        /// </summary>
        public HtmlSelection Selection {
            get {
                if (_selection == null) {
                    _selection = CreateSelection();
                }
                return _selection;
            }
        }

        /// <summary>
        /// The text formatting element of the editor
        /// </summary>
        public HtmlTextFormatting TextFormatting {
            get {
                if (!IsReady) {
                    throw new Exception("HtmlDocument not ready yet!");
                }
                if (textFormatting == null) {
                    textFormatting = new HtmlTextFormatting(this);
                }
                return textFormatting;
            }
        }

        private Interop.IOleUndoManager UndoManager {
            get {
                if (_undoManager == null) {
                    Interop.IServiceProvider serviceProvider = MSHTMLDocument as Interop.IServiceProvider;
                    Debug.Assert(serviceProvider != null);
                    Guid undoManagerGuid = typeof(Interop.IOleUndoManager).GUID;
                    Guid undoManagerGuid2 = typeof(Interop.IOleUndoManager).GUID;
                    IntPtr undoManagerPtr = Interop.NullIntPtr;
                    int hr = serviceProvider.QueryService(ref undoManagerGuid2, ref undoManagerGuid, out undoManagerPtr);
                    if ((hr == Interop.S_OK) && (undoManagerPtr != Interop.NullIntPtr)) {
                        _undoManager = (Interop.IOleUndoManager)Marshal.GetObjectForIUnknown(undoManagerPtr);
                        Marshal.Release(undoManagerPtr);
                    }
                }
                return _undoManager;
            }
        }

        /// <summary>
        /// </summary>
        public void ClearDirtyState() {
            if (!IsReady)
                return;
            Exec(Interop.IDM_SETDIRTY,false);
        }

        /// <summary>
        /// </summary>
        protected virtual HtmlSelection CreateSelection() {
            return new HtmlSelection(this);
        }

        protected override void OnAfterSave() {
            // In non-full document mode, we don't actually save the IPersistInitStream, so
            // clear the dirty bit here
            if (!IsFullDocumentMode) {
                ClearDirtyState();
            }
        }

        /// <summary>
        /// Overridden to remove the grid behavior before loading
        /// </summary>
        protected override void OnBeforeLoad() {
            //are already visible
            if (BordersVisible) {
                BordersVisible = false;
                _bordersDesired = true;
            }
        }
        /// <summary>
        /// Overridden to activate design and multiple selection modes
        /// </summary>
        /// <param name="args"></param>
        protected override void OnCreated(EventArgs args) {
            if (args == null) {
                throw new ArgumentNullException("You must specify a non-null EventArgs for OnCreated");
            }

            base.OnCreated(args);

            object[] mshtmlArgs = new object[1];

            mshtmlArgs[0] = true;
            CommandTarget.Exec(ref Interop.Guid_MSHTML, Interop.IDM_PERSISTDEFAULTVALUES, 0, mshtmlArgs, null);

            mshtmlArgs[0] = true;
            CommandTarget.Exec(ref Interop.Guid_MSHTML, Interop.IDM_PROTECTMETATAGS, 0, mshtmlArgs, null);

            mshtmlArgs[0] = true;
            CommandTarget.Exec(ref Interop.Guid_MSHTML, Interop.IDM_PRESERVEUNDOALWAYS, 0, mshtmlArgs, null);

            mshtmlArgs[0] = true;
            CommandTarget.Exec(ref Interop.Guid_MSHTML, Interop.IDM_NOACTIVATENORMALOLECONTROLS, 0, mshtmlArgs, null);

            mshtmlArgs[0] = true;
            CommandTarget.Exec(ref Interop.Guid_MSHTML, Interop.IDM_NOACTIVATEDESIGNTIMECONTROLS, 0, mshtmlArgs, null);

            mshtmlArgs[0] = true;
            CommandTarget.Exec(ref Interop.Guid_MSHTML, Interop.IDM_NOACTIVATEJAVAAPPLETS, 0, mshtmlArgs, null);

            mshtmlArgs[0] = true;
            CommandTarget.Exec(ref Interop.Guid_MSHTML, Interop.IDM_NOFIXUPURLSONPASTE, 0, mshtmlArgs, null);

            //Set the design mode to the last desired design mode
            if (_designModeDesired) {
                DesignModeEnabled = _designModeDesired;
                _designModeDesired = false;
            }

        }

        /// <summary>
        /// Overridden to set the design mode to the last desired design mode and multiple selection flag
        /// to the last desired multiple selection flag
        /// </summary>
        /// <param name="args"></param>
        protected override internal void OnReadyStateComplete(EventArgs args) {
            base.OnReadyStateComplete(args);

            _persistStream = (Interop.IPersistStreamInit)MSHTMLDocument;

            Selection.SynchronizeSelection();

            //Set the mutiple selection mode to the last desired multiple selection mode
            if (_multipleSelectionDesired) {
                MultipleSelectionEnabled = _multipleSelectionDesired;
            }

            //Set the absolute positioning mode to the last desired absolute position mode
            if (_absolutePositioningDesired) {
                AbsolutePositioningEnabled = _absolutePositioningDesired;
            }

            //Set the absolute positioning mode to the last desired absolute position mode
            if (_bordersDesired) {
                BordersVisible = _bordersDesired;
            }

        }
    }
}
