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
    using System.Diagnostics;
    using System.Collections;
    using System.Runtime.InteropServices;

    /// <summary>
    /// Summary description for HtmlSelection.
    /// </summary>
    public class HtmlSelection {

        public static readonly string DesignTimeLockAttribute = "Design_Time_Lock";

        private EventHandler _selectionChangedHandler;

        private HtmlEditor _editor;
        private Interop.IHTMLDocument2 _document;
        private HtmlSelectionType _type;
        private int _selectionLength;
        private string _text;
        private object _mshtmlSelection;
        private ArrayList _items;
        private ArrayList _elements;
        private bool _sameParentValid;
        private int _maxZIndex;

        public HtmlSelection(HtmlEditor editor) {
            _editor = editor;
            _maxZIndex = 99;
        }

        /// <summary>
        /// Indicates if the current selection can be aligned
        /// </summary>
        public bool CanAlign {
            get {
                if (_items.Count < 2) {
                    return false;
                }
                if (_type == HtmlSelectionType.ElementSelection) {
                    foreach (Interop.IHTMLElement elem in _items) {
                        //First check if they are all absolutely positioned
                        if (!IsElement2DPositioned(elem)) {
                            return false;
                        }

                        //Then check if none of them are locked
                        if (IsElementLocked(elem)) {
                            return false;
                        }
                    }
                    //Then check if they all have the same parent
                    if (!SameParent) {
                        return false;
                    }
                    return true;
                }
                return false;
            }
        }

        /// <summary>
        /// Indicates if the current selection can be size-matched
        /// </summary>
        public bool CanMatchSize {
            get {
                if (_items.Count < 2) {
                    return false;
                }
                if (_type == HtmlSelectionType.ElementSelection) {
                    foreach (Interop.IHTMLElement elem in _items) {
                        //Then check if none of them are locked
                        if (IsElementLocked(elem)) {
                            return false;
                        }
                    }
                    //Then check if they all have the same parent
                    if (!SameParent) {
                        return false;
                    }
                    return true;
                }
                return false;
            }
        }

        /// <summary>
        /// Indicates if the current selection can have it's hyperlink removed
        /// </summary>
        public bool CanRemoveHyperlink {
            get {
                return _editor.IsCommandEnabled(Interop.IDM_UNLINK);
            }
        }

        /// <summary>
        /// Indicates if the current selection have it's z-index modified
        /// </summary>
        public bool CanChangeZIndex {
            get {
                if (_items.Count == 0) {
                    return false;
                }
                if (_type == HtmlSelectionType.ElementSelection) {
                    foreach (Interop.IHTMLElement elem in _items) {
                        //First check if they are all absolutely positioned
                        if (!IsElement2DPositioned(elem)) {
                            return false;
                        }
                    }
                    //Then check if they all have the same parent
                    if (!SameParent) {
                        return false;
                    }
                    return true;
                }
                return false;
            }
        }

        /// <summary>
        /// Indicates if the current selection can be wrapped in HTML tags
        /// </summary>
        public bool CanWrapSelection {
            get {
                if ((_selectionLength != 0) && (Type == HtmlSelectionType.TextSelection)) {
                    return true;
                }
                return false;
            }
        }

        protected HtmlEditor Editor {
            get {
                return _editor;
            }
        }

        public ICollection Elements {
            get {
                if (_elements == null) {
                    _elements = new ArrayList();
                    foreach (Interop.IHTMLElement element in _items) {
                        object wrapper = CreateElementWrapper(element);
                        if (wrapper != null) {
                            _elements.Add(wrapper);
                        }
                    }
                }
                return _elements;
            }
        }

        /// <summary>
        /// All selected items
        /// </summary>
        internal ICollection Items {
            get {
                return _items;
            }
        }

        public int Length {
            get {
                return _selectionLength;
            }
        }

        /// <summary>
        /// Indicates if all items in the selection have the same parent element
        /// </summary>
        private bool SameParent {
            get {
                if (!_sameParentValid) {
                    IntPtr primaryParentElementPtr = Interop.NullIntPtr;

                    foreach (Interop.IHTMLElement elem in _items) {
                        //Check if all items have the same parent by doing pointer equality
                        Interop.IHTMLElement parentElement = elem.GetParentElement();
                        IntPtr parentElementPtr = Marshal.GetIUnknownForObject(parentElement);
                        //If we haven't gotten a primary parent element (ie, this is the first time through the loop)
                        //Remember what the this parent element is
                        if (primaryParentElementPtr == Interop.NullIntPtr) {
                            primaryParentElementPtr = parentElementPtr;
                        }
                        else {
                            //Check the pointers
                            if (primaryParentElementPtr != parentElementPtr) {
                                Marshal.Release(parentElementPtr);
                                if (primaryParentElementPtr != Interop.NullIntPtr) {
                                    Marshal.Release(primaryParentElementPtr);
                                }
                                _sameParentValid = false;
                                return _sameParentValid;
                            }
                            Marshal.Release(parentElementPtr);
                        }
                    }
                    if (primaryParentElementPtr != Interop.NullIntPtr) {
                        Marshal.Release(primaryParentElementPtr);
                    }
                    _sameParentValid = true;
                }
                return _sameParentValid;
            }
        }

        public event EventHandler SelectionChanged {
            add {
                _selectionChangedHandler = (EventHandler)Delegate.Combine(_selectionChangedHandler, value);
            }
            remove {
                if (_selectionChangedHandler != null) {
                    _selectionChangedHandler = (EventHandler)Delegate.Remove(_selectionChangedHandler, value);
                }
            }
        }

        /// <summary>
        /// Returns the MSHTML selection object (IHTMLTxtRange or IHTMLControlRange)
        /// Does not synchronize the selection!!!  Uses the selection from the last synchronization
        /// </summary>
        protected internal object MSHTMLSelection {
            get {
                return _mshtmlSelection;
            }
        }

        /// <summary>
        /// Returns the text contained in the selection if there is a text selection
        /// </summary>
        public string Text {
            get {
                if (_type == HtmlSelectionType.TextSelection) {
                    return _text;
                }
                return null;
            }
        }

        /// <summary>
        /// The HtmlSelectionType of the selection
        /// </summary>
        public HtmlSelectionType Type {
            get {
                return _type;
            }
        }

        public void ClearSelection() {
            _editor.Exec(Interop.IDM_CLEARSELECTION);
        }

        protected virtual object CreateElementWrapper(Interop.IHTMLElement element) {
            return Element.GetWrapperFor(element, _editor);
        }

        /// <summary>
        /// Returns info about the absolute positioning of the selection
        /// </summary>
        /// <returns></returns>
        public HtmlCommandInfo GetAbsolutePositionInfo() {
            return _editor.GetCommandInfo(Interop.IDM_ABSOLUTE_POSITION);
        }

        /// <summary>
        /// Returns info about the design time lock state of the selection
        /// </summary>
        /// <returns></returns>
        public HtmlCommandInfo GetLockInfo() {
            if (_type == HtmlSelectionType.ElementSelection) {
                foreach (Interop.IHTMLElement elem in _items) {
                    //We only need to check that all elements are absolutely positioned
                    if (!IsElement2DPositioned(elem)) {
                        return (HtmlCommandInfo)0;
                    }

                    if (IsElementLocked(elem)) {
                        return HtmlCommandInfo.Checked | HtmlCommandInfo.Enabled;
                    }
                    return HtmlCommandInfo.Enabled;
                }
            }
            return (HtmlCommandInfo)0;
        }

        public string GetOuterHtml() {
            Debug.Assert(Items.Count == 1, "Can't get OuterHtml of more than one element");

            string outerHtml = String.Empty;
            try {
                outerHtml = ((Interop.IHTMLElement)_items[0]).GetOuterHTML();

                // Call this twice because, in the first call, Trident will call OnContentSave, which calls SetInnerHtml, but
                // the outer HTML it returns does not include that new inner HTML.
                outerHtml = ((Interop.IHTMLElement)_items[0]).GetOuterHTML();
            }
            catch {
            }

            return outerHtml;
        }

        public ArrayList GetParentHierarchy(object o) {
            Interop.IHTMLElement current = GetIHtmlElement(o);
            if (current == null) {
                return null;
            }

            string tagName = current.GetTagName().ToLower();
            if (tagName.Equals("body")) {
                return null;
            }

            ArrayList ancestors = new ArrayList();

            current = current.GetParentElement();
            while ((current != null) && (current.GetTagName().ToLower().Equals("body") == false)) {
                Element element = Element.GetWrapperFor(current, _editor);
                if (IsSelectableElement(element)) {
                    ancestors.Add(element);
                }
                current = current.GetParentElement();
            }

            // Don't add the body tag to the hierarchy if we aren't in full document mode
            if (current != null) {
                Element element = Element.GetWrapperFor(current, _editor);
                if (IsSelectableElement(element)) {
                    ancestors.Add(element);
                }
            }

            return ancestors;
        }

        protected virtual Interop.IHTMLElement GetIHtmlElement(object o) {
            if (o is Element) {
                return ((Element)o).Peer;
            }
            return null;
        }

        /// <summary>
        /// Convenience method for checking if the specified element is absolutely positioned
        /// </summary>
        /// <param name="elem"></param>
        /// <returns></returns>
        private bool IsElement2DPositioned(Interop.IHTMLElement elem) {
            Interop.IHTMLElement2 elem2 = (Interop.IHTMLElement2) elem;
            Interop.IHTMLCurrentStyle style = elem2.GetCurrentStyle();
            string position = style.GetPosition();
            if ((position == null) || (String.Compare(position, "absolute", true) != 0)) {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Convenience method for checking if the specified element has a design time lock
        /// </summary>
        /// <param name="elem"></param>
        /// <returns></returns>
        private bool IsElementLocked(Interop.IHTMLElement elem) {
            object[] attribute = new object[1];
            elem.GetAttribute(DesignTimeLockAttribute,0,attribute);
            if (attribute[0] == null) {
                Interop.IHTMLStyle style = elem.GetStyle();
                attribute[0] = style.GetAttribute(DesignTimeLockAttribute,0);
            }
            if ((attribute[0] == null) || !(attribute[0] is string)) {
                return false;
            }
            return true;
        }

        protected virtual bool IsSelectableElement(Element element) {
            return (_editor.IsFullDocumentMode || (element.TagName.ToLower() != "body"));
        }

        protected virtual void OnSelectionChanged() {
            if (_selectionChangedHandler != null) {
                _selectionChangedHandler.Invoke(this, EventArgs.Empty);
            }
        }

        public bool SelectElement(object o) {
            ArrayList list = new ArrayList(1);
            list.Add(o);
            return SelectElements(list);
        }

        public bool SelectElements(ICollection elements) {
            Interop.IHTMLElement body = _editor.MSHTMLDocument.GetBody();
            Interop.IHTMLTextContainer container = body as Interop.IHTMLTextContainer;
            Debug.Assert(container != null);
            object controlRange = container.createControlRange();

            Interop.IHtmlControlRange htmlControlRange = controlRange as Interop.IHtmlControlRange;
            Debug.Assert(htmlControlRange != null);
            if (htmlControlRange == null) {
                return false;
            }

            Interop.IHtmlControlRange2 htmlControlRange2 = controlRange as Interop.IHtmlControlRange2;
            Debug.Assert(htmlControlRange2 != null);
            if (htmlControlRange2 == null) {
                return false;
            }


            int hr = 0;
            foreach (object o in elements) {
                Interop.IHTMLElement element = GetIHtmlElement(o);
                if (element == null) {
                    return false;
                }
                hr = htmlControlRange2.addElement(element);
                if (hr != Interop.S_OK) {
                    break;
                }
            }
            if (hr == Interop.S_OK) {
                //If it succeeded, simply select the control range
                htmlControlRange.Select();
            }
            else {
                // elements like DIV and SPAN, w/o layout, cannot be added to a control selelction.
                Interop.IHtmlBodyElement bodyElement = (Interop.IHtmlBodyElement)body;
                Interop.IHTMLTxtRange textRange = bodyElement.createTextRange();
                if (textRange != null) {
                    foreach (object o in elements) {
                        try {
                            Interop.IHTMLElement element = GetIHtmlElement(o);
                            if (element == null) {
                                return false;
                            }
                            textRange.MoveToElementText(element);
                        }
                        catch {
                        }
                    }
                    textRange.Select();
                }
            }
            return true;
        }

        //        /// <summary>
        //        /// Sends all selected items to the back
        //        /// </summary>
        //        public void SendToBack() {
        //            //TODO: How do we compress the ZIndexes so they never go out of the range of an int
        //            SynchronizeSelection();
        //            if (_type == HtmlSelectionType.ElementSelection) {
        //                if (_items.Count > 1) {
        //                    //We have to move all items to the back, and maintain their ordering, so
        //                    //Find the maximum ZIndex in the group
        //                    int max = _minZIndex;
        //                    int count = _items.Count;
        //                    Interop.IHTMLStyle[] styles = new Interop.IHTMLStyle[count];
        //                    int[] zIndexes = new int[count];
        //                    for (int i = 0; i < count; i++) {
        //                        Interop.IHTMLElement elem = (Interop.IHTMLElement)_items[i];
        //                        styles[i] = elem.GetStyle();
        //                        zIndexes[i] = (int)styles[i].GetZIndex();
        //                        if (zIndexes[i] > max) {
        //                            max = zIndexes[i];
        //                        }
        //                    }
        //                    //Calculate how far the first element has to be moved in order to be in the back
        //                    int offset = max - (_minZIndex - 1);
        //                    BatchedUndoUnit unit = _editor.OpenBatchUndo("Align Left");
        //                    try {
        //                        //Then send all items in the selection that far back
        //                        for (int i = 0; i < count; i++) {
        //                            int newPos = zIndexes[i] - offset;
        //                            if (zIndexes[i] == _maxZIndex) {
        //                                _maxZIndex--;
        //                            }
        //                            styles[i].SetZIndex(newPos);
        //                            if (newPos < _minZIndex) {
        //                                _minZIndex = newPos;
        //                            }
        //                        }
        //                    }
        //                    catch (Exception e) {
        //                        System.Windows.Forms.MessageBox.Show(e.ToString(),"Exception");
        //                    }
        //                    finally {
        //                        unit.Close();
        //                    }
        //                }
        //                else {
        //                    Interop.IHTMLElement elem = (Interop.IHTMLElement)_items[0];
        //                    object zIndex = elem.GetStyle().GetZIndex();
        //                    if ((zIndex != null) && !(zIndex is DBNull)) {
        //                        if ((int)zIndex == _minZIndex) {
        //                            // if the element is already in the back do nothing.
        //                            return;
        //                        }
        //
        //                        if ((int)zIndex == _maxZIndex) {
        //                            _maxZIndex--;
        //                        }
        //                    }
        //                    elem.GetStyle().SetZIndex(--_minZIndex);
        //                }
        //            }
        //        }

        public void SetOuterHtml(string outerHtml) {
            Debug.Assert(Items.Count == 1, "Can't get OuterHtml of more than one element");
            ((Interop.IHTMLElement)_items[0]).SetOuterHTML(outerHtml);
        }

        /// <summary>
        /// Synchronizes the selection state held in this object with the selection state in MSHTML
        /// </summary>
        /// <returns>true if the selection has changed</returns>
        public bool SynchronizeSelection() {
            //Get the selection object from the MSHTML document
            if (_document == null) {
                _document = _editor.MSHTMLDocument;
            }
            Interop.IHTMLSelectionObject selectionObj = _document.GetSelection();

            //Get the current selection from that selection object
            object currentSelection = null;
            try {
                currentSelection = selectionObj.CreateRange();
            }
            catch {
            }

            ArrayList oldItems = _items;
            HtmlSelectionType oldType = _type;
            int oldLength = _selectionLength;
            //Default to an empty selection
            _type = HtmlSelectionType.Empty;
            _selectionLength = 0;
            if (currentSelection != null) {
                _mshtmlSelection = currentSelection;
                _items = new ArrayList();
                //If it's a text selection
                if (currentSelection is Interop.IHTMLTxtRange) {
                    Interop.IHTMLTxtRange textRange = (Interop.IHTMLTxtRange) currentSelection;
                    //IntPtr ptr = Marshal.GetIUnknownForObject(textRange);
                    Interop.IHTMLElement parentElement = textRange.ParentElement();
                    // If the document is in full document mode or we're selecting a non-body tag, allow it to select
                    // otherwise, leave the selection as empty (since we don't want the body tag to be selectable on an ASP.NET
                    // User Control
                    if (IsSelectableElement(Element.GetWrapperFor(parentElement, _editor))) {
                        //Add the parent of the text selection
                        if (parentElement != null) {
                            _text = textRange.GetText();
                            if (_text != null) {
                                _selectionLength = _text.Length;
                            }
                            else {
                                _selectionLength = 0;
                            }
                            _type = HtmlSelectionType.TextSelection;
                            _items.Add(parentElement);
                        }
                    }
                }
                    //If it's a control selection
                else if (currentSelection is Interop.IHtmlControlRange) {
                    Interop.IHtmlControlRange controlRange = (Interop.IHtmlControlRange) currentSelection;
                    int selectedCount = controlRange.GetLength();
                    //Add all elements selected
                    if (selectedCount > 0) {
                        _type = HtmlSelectionType.ElementSelection;
                        for (int i = 0; i < selectedCount; i++) {
                            Interop.IHTMLElement currentElement = controlRange.Item(i);
                            _items.Add(currentElement);
                        }
                        _selectionLength = selectedCount;
                    }
                }
            }
            _sameParentValid = false;

            bool selectionChanged = false;
            //Now check if there was a change of selection
            //If the two selections have different lengths, then the selection has changed
            if (_type != oldType) {
                selectionChanged = true;
            }
            else if (_selectionLength != oldLength) {
                selectionChanged = true;
            }
            else {
                if (_items != null) {
                    //If the two selections have a different element, then the selection has changed
                    for (int i = 0; i < _items.Count; i++) {
                        if (_items[i] != oldItems[i]) {
                            selectionChanged = true;
                            break;
                        }
                    }
                }
            }
            if (selectionChanged) {
                //Set _elements to null so no one can retrieve a dirty copy of the selection element wrappers
                _elements = null;

                OnSelectionChanged();
                return true;
            }

            return false;
        }
        /// <summary>
        /// Toggle the absolute positioning state of the selected items
        /// </summary>
        public void ToggleAbsolutePosition() {
            _editor.Exec(Interop.IDM_ABSOLUTE_POSITION, !((GetAbsolutePositionInfo() & HtmlCommandInfo.Checked) != 0));
            SynchronizeSelection();
            if (_type == HtmlSelectionType.ElementSelection) {
                foreach (Interop.IHTMLElement elem in _items) {
                    elem.GetStyle().SetZIndex(++_maxZIndex);
                }
            }
        }

        /// <summary>
        /// Toggle the design time lock state of the selected items
        /// </summary>
        public void ToggleLock() {
            //Switch the lock on each item
            foreach (Interop.IHTMLElement elem in _items) {
                Interop.IHTMLStyle style = elem.GetStyle();
                if (IsElementLocked(elem)) {
                    //We need to remove attributes off the element and the style because of a bug in Trident
                    elem.RemoveAttribute(DesignTimeLockAttribute,0);
                    style.RemoveAttribute(DesignTimeLockAttribute,0);
                }
                else {
                    //We need to add attributes to the element and the style because of a bug in Trident
                    elem.SetAttribute(DesignTimeLockAttribute,"true",0);
                    style.SetAttribute(DesignTimeLockAttribute,"true",0);
                }
            }
        }

        public void WrapSelection(string tag) {
            WrapSelection(tag, null);
        }

        public void WrapSelection(string tag, IDictionary attributes) {
            //Create a string for all the attributes
            string attributeString = String.Empty;
            if (attributes != null) {
                foreach (string key in attributes.Keys) {
                    attributeString+=key+"=\""+attributes[key]+"\" ";
                }
            }
            SynchronizeSelection();
            if (_type == HtmlSelectionType.TextSelection) {
                Interop.IHTMLTxtRange textRange = (Interop.IHTMLTxtRange)MSHTMLSelection;
                string oldText = textRange.GetHtmlText();
                if (oldText == null) {
                    oldText = String.Empty;
                }
                string newText = "<"+tag+" "+attributeString+">"+oldText+"</"+tag+">";
                textRange.PasteHTML(newText);
            }
        }

        public void WrapSelectionInDiv() {
            WrapSelection("div");
        }

        public void WrapSelectionInSpan() {
            WrapSelection("span");
        }

        public void WrapSelectionInBlockQuote() {
            WrapSelection("blockquote");
        }

        public void WrapSelectionInHyperlink(string url) {
            _editor.Exec(Interop.IDM_HYPERLINK,url);
        }

        public void RemoveHyperlink() {
            _editor.Exec(Interop.IDM_UNLINK);
        }
    }
}
