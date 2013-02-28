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
    using System.Globalization;
    using System.Collections;
    using System.Reflection;

    /// <summary>
    /// The base class for all element wrappers. These provide information for populating
    /// the property grid.
    /// </summary>
    [
    DesignOnly(true)
    ]
    public class Element {
        public static Element GetWrapperFor(Interop.IHTMLElement element, HtmlControl owner) {
            Element wrapperElement = new Element(element);
            wrapperElement.SetOwner(owner);
            return wrapperElement;
        }

        private Interop.IHTMLElement _peer;
        private HtmlControl _owner;

        internal Element(Interop.IHTMLElement peer) {
            Debug.Assert(peer != null);
            _peer = peer;
        }

        [Browsable(false)]
        public string InnerHtml {
            get {
                try {
                    return _peer.GetInnerHTML();
                }
                catch (Exception e) {
                    Debug.Fail(e.ToString(), "Could not get Element InnerHTML");

                    return String.Empty;
                }
            }
            set {
                try {
                    _peer.SetInnerHTML(value);
                }
                catch (Exception e) {
                    Debug.Fail(e.ToString(), "Could not set Element InnerHTML");
                }
            }
        }

        [Browsable(false)]
        public string OuterHtml {
            get {
                try {
                    return _peer.GetOuterHTML();
                }
                catch (Exception e) {
                    Debug.Fail(e.ToString(), "Could not get Element OuterHTML");

                    return String.Empty;
                }
            }
            set {
                try {
                    _peer.SetOuterHTML(value);
                }
                catch (Exception e) {
                    Debug.Fail(e.ToString(), "Could not set Element OuterHTML");
                }
            }
        }

        [Browsable(false)]
        public string TagName {
            get {
                try {
                    return _peer.GetTagName();
                }
                catch (Exception e) {
                    Debug.Fail(e.ToString(), "Could not get Element TagName" + e.ToString());

                    return String.Empty;
                }
            }
        }

        internal Interop.IHTMLElement Peer {
            get {
                return _peer;
            }
        }

        public object GetAttribute(string attribute) {
            try {
                object[] obj = new object[1];

                _peer.GetAttribute(attribute, 0, obj);

                object o = obj[0];
                if (o is DBNull) {
                    o = null;
                }
                return o;
            }
            catch (Exception e) {
                Debug.Fail(e.ToString(), "Call to IHTMLElement::GetAttribute failed in Element");

                return null;
            }
        }

        protected internal bool GetBooleanAttribute(string attribute) {
            object o = GetAttribute(attribute);

            if (o == null) {
                return false;
            }

            Debug.Assert(o is bool, "Attribute " + attribute + " is not of type Boolean");
            if (o is bool) {
                return (bool)o;
            }

            return false;
        }

        protected internal Color GetColorAttribute(string attribute) {
            string color = GetStringAttribute(attribute);
            
            if (color.Length == 0) {
                return Color.Empty;
            }
            else {
                return ColorTranslator.FromHtml(color);
            }
        }

        protected internal Enum GetEnumAttribute(string attribute, Enum defaultValue) {
            Type enumType = defaultValue.GetType();

            object o = GetAttribute(attribute);
            if (o == null) {
                return defaultValue;
            }

            Debug.Assert(o is string, "Attribute " + attribute + " is not of type String");
            string s = o as string;
            if ((s == null) || (s.Length == 0)) {
                return defaultValue;
            }

            try {
                return (Enum)Enum.Parse(enumType, s, true);
            }
            catch {
                return defaultValue;
            }
        }

        protected internal int GetIntegerAttribute(string attribute, int defaultValue) {
            object o = GetAttribute(attribute);

            if (o == null) {
                return defaultValue;
            }
            if (o is int) {
                return (int)o;
            }
            if (o is short) {
                return (short)o;
            }
            if (o is string) {
                string s = (string)o;
                if ((s.Length != 0) && (Char.IsDigit(s[0]))) {
                    try {
                        return Int32.Parse((string)o);
                    }
                    catch {
                    }
                }
            }

            Debug.Fail("Attribute " + attribute + " is not an integer");
            return defaultValue;
        }
        
        public Element GetChild(int index) {
            Interop.IHTMLElementCollection children = (Interop.IHTMLElementCollection)_peer.GetChildren();
            Interop.IHTMLElement child = (Interop.IHTMLElement)children.Item(null, index);

            return Element.GetWrapperFor(child, _owner);
        }

        public Element GetChild(string name) {
            Interop.IHTMLElementCollection children = (Interop.IHTMLElementCollection)_peer.GetChildren();
            Interop.IHTMLElement child = (Interop.IHTMLElement)children.Item(name, null);

            return Element.GetWrapperFor(child, _owner);
        }

        public Element GetParent() {
            Interop.IHTMLElement parent = (Interop.IHTMLElement)_peer.GetParentElement();
            return Element.GetWrapperFor(parent, _owner);
        }

        protected string GetRelativeUrl(string absoluteUrl) {
            if ((absoluteUrl == null) || (absoluteUrl.Length == 0)) {
                return String.Empty;
            }
            
            string s = absoluteUrl;
            if (_owner != null) {
                string ownerUrl = _owner.Url;

                if (ownerUrl.Length != 0) {
                    try {
                        Uri ownerUri = new Uri(ownerUrl);
                        Uri imageUri = new Uri(s);

                        s = ownerUri.MakeRelative(imageUri);
                    }
                    catch {
                    }
                }
            }

            return s;
        }

        protected internal string GetStringAttribute(string attribute) {
            return GetStringAttribute(attribute, String.Empty);
        }

        protected internal string GetStringAttribute(string attribute, string defaultValue) {
            object o = GetAttribute(attribute);

            if (o == null) {
                return defaultValue;
            }
            if (o is string) {
                return (string)o;
            }

            return defaultValue;
        }

        public void RemoveAttribute(string attribute) {
            try {
                _peer.RemoveAttribute(attribute, 0);
            }
            catch (Exception e) {
                Debug.Fail(e.ToString(), "Call to IHTMLElement::RemoveAttribute failed in Element");
            }
        }

        public void SetAttribute(string attribute, object value) {
            try {
                _peer.SetAttribute(attribute, value, 0);
            }
            catch (Exception e) {
                Debug.Fail(e.ToString(), "Call to IHTMLElement::SetAttribute failed in Element");
            }
        }

        protected internal void SetBooleanAttribute(string attribute, bool value) {
            if (value) {
                SetAttribute(attribute, true);
            }
            else {
                RemoveAttribute(attribute);
            }
        }

        protected internal void SetColorAttribute(string attribute, Color value) {
            if (value.IsEmpty) {
                RemoveAttribute(attribute);
            }
            else {
                SetAttribute(attribute, ColorTranslator.ToHtml(value));
            }
        }

        protected internal void SetEnumAttribute(string attribute, Enum value, Enum defaultValue) {
            Debug.Assert(value.GetType().Equals(defaultValue.GetType()));

            if (value.Equals(defaultValue)) {
                RemoveAttribute(attribute);
            }
            else {
                SetAttribute(attribute, value.ToString(CultureInfo.InvariantCulture));
            }
        }

        protected internal void SetIntegerAttribute(string attribute, int value, int defaultValue) {
            if (value == defaultValue) {
                RemoveAttribute(attribute);
            }
            else {
                SetAttribute(attribute, value);
            }
        }

        internal void SetOwner(HtmlControl owner) {
            _owner = owner;
        }

        protected internal void SetStringAttribute(string attribute, string value) {
            SetStringAttribute(attribute, value, String.Empty);
        }

        protected internal void SetStringAttribute(string attribute, string value, string defaultValue) {
            if ((value == null) || value.Equals(defaultValue)) {
                RemoveAttribute(attribute);
            }
            else {
                SetAttribute(attribute, value);
            }
        }

        public override string ToString() {
            if (_peer != null) {
                try {
                    return "<" + _peer.GetTagName() + ">";
                }
                catch {
                }
            }
            return String.Empty;
        }
    }
}
