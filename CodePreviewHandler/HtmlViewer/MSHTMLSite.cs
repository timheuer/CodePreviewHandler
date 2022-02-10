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
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    [
    ClassInterface(ClassInterfaceType.None)
    ]
    internal class MSHTMLSite :
        Interop.IOleClientSite,
        Interop.IOleContainer,
        Interop.IOleDocumentSite,
        Interop.IOleInPlaceSite,
        Interop.IOleInPlaceFrame,
        Interop.IDocHostUIHandler,
        Interop.IPropertyNotifySink,
        Interop.IAdviseSink,
        Interop.IOleServiceProvider {

        /// the Control used to host (and parent) the mshtml window
        private HtmlControl hostControl;

        /// the mshtml instance and various related objects
        private Interop.IOleObject tridentOleObject;
        private Interop.IHTMLDocument2 tridentDocument;
        private Interop.IOleCommandTarget tridentCmdTarget;
        private Interop.IOleDocumentView tridentView;
        private Interop.IOleInPlaceActiveObject activeObject;

        // cookie representing our sink
        private Interop.ConnectionPointCookie propNotifyCookie;
        private int adviseSinkCookie;

        //        private DropTarget _dropTarget;

        /// <summary>
        /// </summary>
        public MSHTMLSite(HtmlControl hostControl) {
            if ((hostControl == null) || (hostControl.IsHandleCreated == false)) {
                throw new ArgumentException();
            }

            this.hostControl = hostControl;
            hostControl.Resize += new EventHandler(this.OnParentResize);
        }

        /// <summary>
        /// </summary>
        public Interop.IOleCommandTarget MSHTMLCommandTarget {
            get {
                return tridentCmdTarget;
            }
        }

        /// <summary>
        /// </summary>
        public Interop.IHTMLDocument2 MSHTMLDocument {
            get {
                return tridentDocument;
            }
        }

        /// <summary>
        /// </summary>
        public void ActivateMSHTML() {
            Debug.Assert(tridentOleObject != null, "How'd we get here when trident is null!");

            try {
                Interop.COMRECT r = new Interop.COMRECT();

                Interop.GetClientRect(hostControl.Handle, r);
                tridentOleObject.DoVerb(Interop.OLEIVERB_UIACTIVATE, Interop.NullIntPtr, (Interop.IOleClientSite)this,
                    0, hostControl.Handle, r);
            }
            catch (Exception e) {
                Debug.Fail(e.ToString());
            }
        }

        /// <summary>
        /// </summary>
        public void CloseMSHTML() {
            hostControl.Resize -= new EventHandler(this.OnParentResize);

            try {
                if (propNotifyCookie != null) {
                    propNotifyCookie.Disconnect();
                    propNotifyCookie = null;
                }

                if (tridentDocument != null) {
                    tridentView = null;
                    tridentDocument = null;
                    tridentCmdTarget = null;
                    activeObject = null;

                    if (adviseSinkCookie != 0) {
                        tridentOleObject.Unadvise(adviseSinkCookie);
                        adviseSinkCookie = 0;
                    }

                    tridentOleObject.Close(Interop.OLECLOSE_NOSAVE);
                    tridentOleObject.SetClientSite(null);
                    tridentOleObject = null;
                }
            }
            catch (Exception e) {
                Debug.Fail(e.ToString());
            }
        }

        /// <summary>
        /// </summary>
        public void CreateMSHTML() {
            Debug.Assert(tridentDocument == null, "Must call CloseMSHTML before recreating.");

            bool created = false;
            try {
                // create the trident instance
                tridentDocument = (Interop.IHTMLDocument2)new Interop.HTMLDocument();
                tridentOleObject = (Interop.IOleObject)tridentDocument;

                // hand it our Interop.IOleClientSite implementation
                tridentOleObject.SetClientSite((Interop.IOleClientSite)this);

                created = true;

                propNotifyCookie = new Interop.ConnectionPointCookie(tridentDocument, this, typeof(Interop.IPropertyNotifySink), false);

                tridentOleObject.Advise((Interop.IAdviseSink)this, out adviseSinkCookie);
                Debug.Assert(adviseSinkCookie != 0);

                tridentCmdTarget = (Interop.IOleCommandTarget)tridentDocument;
            }
            finally {
                if (created == false) {
                    tridentDocument = null;
                    tridentOleObject = null;
                    tridentCmdTarget = null;
                }
            }
        }

        /// <summary>
        /// </summary>
        public void DeactivateMSHTML() {
            // TODO: Implement this... once I know how to do it!
        }

        /// <summary>
        /// </summary>
        private void OnParentResize(object src, EventArgs e) {
            if (tridentView != null) {
                Interop.COMRECT r = new Interop.COMRECT();

                Interop.GetClientRect(hostControl.Handle, r);
                tridentView.SetRect(r);
            }
        }

        /// <summary>
        /// </summary>
        private void OnReadyStateChanged() {
            string readyState = tridentDocument.GetReadyState();
            if (String.Compare(readyState, "complete", true) == 0) {
                OnReadyStateComplete();
            }
        }

        /// <summary>
        /// </summary>
        private void OnReadyStateComplete() {
            hostControl.OnReadyStateComplete(EventArgs.Empty);
        }

        internal void SetFocus() {
            if (activeObject != null) {
                IntPtr hWnd = IntPtr.Zero;
                if (activeObject.GetWindow(out hWnd) == Interop.S_OK) {
                    Debug.Assert(hWnd != IntPtr.Zero);
                    Interop.SetFocus(hWnd);
                }
            }
        }

        /// <summary>
        /// </summary>
        public bool TranslateAccelarator(Interop.COMMSG msg) {
            if (activeObject != null) {
                int hr = activeObject.TranslateAccelerator(msg);
                if (hr != Interop.S_FALSE) {
                    return true;
                }
            }
            return false;
        }


        ///////////////////////////////////////////////////////////////////////////
        // Interop.IOleClientSite Implementation

        public int SaveObject() {
            return Interop.S_OK;
        }

        public int GetMoniker(int dwAssign, int dwWhichMoniker, out object ppmk) {
            ppmk = null;
            return Interop.E_NOTIMPL;
        }

        public int GetContainer(out Interop.IOleContainer ppContainer) {
            ppContainer = (Interop.IOleContainer)this;
            return Interop.S_OK;
        }

        public int ShowObject() {
            return Interop.S_OK;
        }

        public int OnShowWindow(int fShow) {
            return Interop.S_OK;
        }

        public int RequestNewObjectLayout() {
            return Interop.S_OK;
        }


        ///////////////////////////////////////////////////////////////////////////
        // Interop.IOleContainer Implementation

        public void ParseDisplayName(object pbc, string pszDisplayName, int[] pchEaten, object[] ppmkOut) {
            Debug.Fail("ParseDisplayName - " + pszDisplayName);
            throw new COMException(String.Empty, Interop.E_NOTIMPL);
        }

        public void EnumObjects(int grfFlags, object[] ppenum) {
            throw new COMException(String.Empty, Interop.E_NOTIMPL);
        }

        public void LockContainer(int fLock) {
        }


        ///////////////////////////////////////////////////////////////////////////
        // Interop.IOleDocumentSite Implementation

        public int ActivateMe(Interop.IOleDocumentView pViewToActivate) {
            Debug.Assert(pViewToActivate != null,
                "Expected the view to be non-null");
            if (pViewToActivate == null)
                return Interop.E_INVALIDARG;

            Interop.COMRECT r = new Interop.COMRECT();

            Interop.GetClientRect(hostControl.Handle, r);

            tridentView = pViewToActivate;
            tridentView.SetInPlaceSite((Interop.IOleInPlaceSite)this);
            tridentView.UIActivate(1);
            tridentView.SetRect(r);
            tridentView.Show(1);

            return Interop.S_OK;
        }


        ///////////////////////////////////////////////////////////////////////////
        // Interop.IOleInPlaceSite Implementation

        public IntPtr GetWindow() {
            return hostControl.Handle;
        }

        public void ContextSensitiveHelp(int fEnterMode) {
            throw new COMException(String.Empty, Interop.E_NOTIMPL);
        }

        public int CanInPlaceActivate() {
            return Interop.S_OK;
        }

        public void OnInPlaceActivate() {
        }

        public void OnUIActivate() {
        }

        public void GetWindowContext(out Interop.IOleInPlaceFrame ppFrame, out Interop.IOleInPlaceUIWindow ppDoc, Interop.COMRECT lprcPosRect, Interop.COMRECT lprcClipRect, Interop.tagOIFI lpFrameInfo) {
            ppFrame = (Interop.IOleInPlaceFrame)this;
            ppDoc = null;

            Interop.GetClientRect(hostControl.Handle, lprcPosRect);
            Interop.GetClientRect(hostControl.Handle, lprcClipRect);

            lpFrameInfo.cb = Marshal.SizeOf(typeof(Interop.tagOIFI));
            lpFrameInfo.fMDIApp = 0;
            lpFrameInfo.hwndFrame = hostControl.Handle;
            lpFrameInfo.hAccel = Interop.NullIntPtr;
            lpFrameInfo.cAccelEntries = 0;
        }

        public int Scroll(Interop.tagSIZE scrollExtant) {
            return Interop.E_NOTIMPL;
        }

        public void OnUIDeactivate(int fUndoable) {
            // NOTE, nikhilko, 7/99: Don't return E_NOTIMPL. Somehow doing nothing and returning S_OK
            //    fixes trident hosting in Win2000.
        }

        public void OnInPlaceDeactivate() {
        }

        public void DiscardUndoState() {
            throw new COMException(String.Empty, Interop.E_NOTIMPL);
        }

        public void DeactivateAndUndo() {
        }

        public int OnPosRectChange(Interop.COMRECT lprcPosRect) {
            return Interop.S_OK;
        }


        ///////////////////////////////////////////////////////////////////////////
        // Interop.IOleInPlaceFrame Implementation

        public void GetBorder(Interop.COMRECT lprectBorder) {
            throw new COMException(String.Empty, Interop.E_NOTIMPL);
        }

        public void RequestBorderSpace(Interop.COMRECT pborderwidths) {
            throw new COMException(String.Empty, Interop.E_NOTIMPL);
        }

        public void SetBorderSpace(Interop.COMRECT pborderwidths) {
            throw new COMException(String.Empty, Interop.E_NOTIMPL);
        }

        public void SetActiveObject(Interop.IOleInPlaceActiveObject pActiveObject, string pszObjName) {
            this.activeObject = pActiveObject;
        }

        public void InsertMenus(IntPtr hmenuShared, Interop.tagOleMenuGroupWidths lpMenuWidths) {
            throw new COMException(String.Empty, Interop.E_NOTIMPL);
        }

        public void SetMenu(IntPtr hmenuShared, IntPtr holemenu, IntPtr hwndActiveObject) {
            throw new COMException(String.Empty, Interop.E_NOTIMPL);
        }

        public void RemoveMenus(IntPtr hmenuShared) {
            throw new COMException(String.Empty, Interop.E_NOTIMPL);
        }

        public void SetStatusText(string pszStatusText) {
        }

        public void EnableModeless(int fEnable) {
        }

        public int TranslateAccelerator(Interop.COMMSG lpmsg, short wID) {
            return Interop.S_FALSE;
        }


        ///////////////////////////////////////////////////////////////////////////
        // IDocHostUIHandler Implementation

        public int ShowContextMenu(int dwID, Interop.POINT pt, object pcmdtReserved, object pdispReserved) {
            //            Point location = hostControl.PointToClient(new Point(pt.x, pt.y));
            //
            //            ShowContextMenuEventArgs e = new ShowContextMenuEventArgs(location, false);
            //
            //            try {
            //                hostControl.OnShowContextMenu(e);
            //            }
            //            catch {
            //                // Make sure we return Interop.S_OK
            //            }
            //
            return Interop.S_OK;
        }

        public int GetHostInfo(Interop.DOCHOSTUIINFO info) {
            info.dwDoubleClick = Interop.DOCHOSTUIDBLCLICK_DEFAULT;
            int flags = 0;
            if (hostControl.AllowInPlaceNavigation) {
                flags |= Interop.DOCHOSTUIFLAG_ENABLE_INPLACE_NAVIGATION;
            }
            if (!hostControl.Border3d) {
                flags |= Interop.DOCHOSTUIFLAG_NO3DBORDER;
            }
            if (!hostControl.ScriptEnabled) {
                flags |= Interop.DOCHOSTUIFLAG_DISABLE_SCRIPT_INACTIVE;
            }
            if (!hostControl.ScrollBarsEnabled) {
                flags |= Interop.DOCHOSTUIFLAG_SCROLL_NO;
            }
            if (hostControl.FlatScrollBars) {
                flags |= Interop.DOCHOSTUIFLAG_FLAT_SCROLLBAR;
            }
            info.dwFlags = flags;
            return Interop.S_OK;
        }

        public int EnableModeless(bool fEnable) {
            return Interop.S_OK;
        }

        public int ShowUI(int dwID, Interop.IOleInPlaceActiveObject activeObject, Interop.IOleCommandTarget commandTarget, Interop.IOleInPlaceFrame frame, Interop.IOleInPlaceUIWindow doc) {
            return Interop.S_OK;
        }

        public int HideUI() {
            return Interop.S_OK;
        }

        public int UpdateUI() {
            return Interop.S_OK;
        }

        public int OnDocWindowActivate(bool fActivate) {
            return Interop.E_NOTIMPL;
        }

        public int OnFrameWindowActivate(bool fActivate) {
            return Interop.E_NOTIMPL;
        }

        public int ResizeBorder(Interop.COMRECT rect, Interop.IOleInPlaceUIWindow doc, bool fFrameWindow) {
            return Interop.E_NOTIMPL;
        }

        public int GetOptionKeyPath(string[] pbstrKey, int dw) {
            pbstrKey[0] = null;
            return Interop.S_OK;
        }

        public int GetDropTarget(Interop.IOleDropTarget pDropTarget, out Interop.IOleDropTarget ppDropTarget) {
            ppDropTarget = null;
            return Interop.E_NOTIMPL;
        }

        public int GetExternal(out object ppDispatch) {
            ppDispatch = hostControl.ScriptObject;
            if (ppDispatch != null) {
                return Interop.S_OK;
            }
            else {
                return Interop.E_NOTIMPL;
            }
        }

        public int TranslateAccelerator(Interop.COMMSG msg, ref Guid group, int nCmdID) {
            return Interop.S_FALSE;
        }

        public int TranslateUrl(int dwTranslate, string strURLIn, out string pstrURLOut) {
            pstrURLOut = null;
            return Interop.E_NOTIMPL;
        }

        public int FilterDataObject(Interop.IOleDataObject pDO, out Interop.IOleDataObject ppDORet) {
            ppDORet = null;
            return Interop.E_NOTIMPL;
        }


        ///////////////////////////////////////////////////////////////////////////
        // IPropertyNotifySink Implementation

        public void OnChanged(int dispID) {
            if (dispID == Interop.DISPID_READYSTATE) {
                OnReadyStateChanged();
            }
        }

        public void OnRequestEdit(int dispID) {
        }


        ///////////////////////////////////////////////////////////////////////////
        // IAdviseSink Implementation

        public void OnDataChange(Interop.FORMATETC pFormat, Interop.STGMEDIUM pStg) {
        }

        public void OnViewChange(int dwAspect, int index) {
        }

        public void OnRename(object pmk) {
        }

        public void OnSave() {
        }

        public void OnClose() {
        }


        ///////////////////////////////////////////////////////////////////////////
        // Interop.IOleServiceProvider

        public int QueryService(ref Guid sid, ref Guid iid, out IntPtr ppvObject) {
            int hr = Interop.E_NOINTERFACE;
            ppvObject = Interop.NullIntPtr;

            //            object service = hostControl.GetService(ref sid);
            //            if (service != null) {
            //                if (iid.Equals(Interop.IID_IUnknown)) {
            //                    ppvObject = Marshal.GetIUnknownForObject(service);
            //                }
            //                else {
            //                    IntPtr pUnk = Marshal.GetIUnknownForObject(service);
            //
            //                    hr = Marshal.QueryInterface(pUnk, ref iid, out ppvObject);
            //                    Marshal.Release(pUnk);
            //                }
            //            }

            return hr;
        }

    }
}

