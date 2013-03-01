using System;
using System.Runtime.InteropServices;
using TimHeuer.ManagedPreviewHandler;

namespace TimHeuer.PreviewHandlers
{
    internal static class PreviewHandlerRegistration
    {
        [ComRegisterFunction]
        private static void Register(Type t) { PreviewHandler.Register(t); }

        [ComUnregisterFunction]
        private static void Unregister(Type t) { PreviewHandler.Unregister(t); }
    }
}
