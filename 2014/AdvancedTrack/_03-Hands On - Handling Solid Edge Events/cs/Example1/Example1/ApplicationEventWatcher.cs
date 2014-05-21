using SolidEdgeContrib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Example1
{
    class ApplicationEventWatcher : ConnectionPointControllerBase, SolidEdgeFramework.ISEApplicationEvents, IDisposable
    {
        private bool _disposed = false;
        private Form1 _form;

        public ApplicationEventWatcher(Form1 form, SolidEdgeFramework.Application application)
        {
            _form = form;

            this.AdviseSink<SolidEdgeFramework.ISEApplicationEvents>(application);
        }

        #region IDisposable implementation

        ~ApplicationEventWatcher()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    this.UnadviseAllSinks();
                }

                _disposed = true;
            }
        }

#endregion

        // Note: Events are fired in a background thread. You cannot update the UI
        // "directly" from a background thread. See ControlExtensions.BeginInvokeIfRequired().
        // Thread.CurrentThread.GetApartmentState() will always be ApartmentState.MTA.
        // OleMessageFilter is not in effect in this thread for two reasons. 1) It's a
        // different thread. 2) It can't be because the ApartmentState = MTA.

        #region SolidEdgeFramework.ISEApplicationEvents

        public void AfterActiveDocumentChange(object theDocument)
        {
            _form.BeginInvokeIfRequired(x =>
            {
                x.OnAfterActiveDocumentChange(theDocument);
            });
        }

        public void AfterCommandRun(int theCommandID)
        {
            _form.BeginInvokeIfRequired(x =>
            {
                x.OnAfterCommandRun(theCommandID);
            });
        }

        public void AfterDocumentOpen(object theDocument)
        {
            _form.BeginInvokeIfRequired(x =>
            {
                x.OnAfterDocumentOpen(theDocument);
            });
        }

        public void AfterDocumentPrint(object theDocument, int hDC, ref double ModelToDC, ref int Rect)
        {
            // Cannot use ref or out parameter 'ModelToDC' inside an anonymous method, lambda expression, or query expression.
            // Cannot use ref or out parameter 'Rect' inside an anonymous method, lambda expression, or query expression.
            var a = ModelToDC;
            var b = Rect;

            _form.InvokeIfRequired(x =>
            {
                x.OnAfterDocumentPrint(theDocument, hDC, a, b);
            });
        }

        public void AfterDocumentSave(object theDocument)
        {
            _form.BeginInvokeIfRequired(x =>
            {
                x.OnAfterDocumentSave(theDocument);
            });
        }

        public void AfterEnvironmentActivate(object theEnvironment)
        {
            _form.BeginInvokeIfRequired(x =>
            {
                x.OnAfterEnvironmentActivate(theEnvironment);
            });
        }

        public void AfterNewDocumentOpen(object theDocument)
        {
            _form.BeginInvokeIfRequired(x =>
            {
                x.OnAfterNewDocumentOpen(theDocument);
            });
        }

        public void AfterNewWindow(object theWindow)
        {
            _form.BeginInvokeIfRequired(x =>
            {
                x.OnAfterNewWindow(theWindow);
            });
        }

        public void AfterWindowActivate(object theWindow)
        {
            _form.BeginInvokeIfRequired(x =>
            {
                x.OnAfterWindowActivate(theWindow);
            });
        }

        public void BeforeCommandRun(int theCommandID)
        {
            _form.BeginInvokeIfRequired(x =>
            {
                x.OnBeforeCommandRun(theCommandID);
            });
        }

        public void BeforeDocumentClose(object theDocument)
        {
            _form.BeginInvokeIfRequired(x =>
            {
                x.OnBeforeDocumentClose(theDocument);
            });
        }

        public void BeforeDocumentPrint(object theDocument, int hDC, ref double ModelToDC, ref int Rect)
        {
            // Cannot use ref or out parameter 'ModelToDC' inside an anonymous method, lambda expression, or query expression.
            // Cannot use ref or out parameter 'Rect' inside an anonymous method, lambda expression, or query expression.
            var a = ModelToDC;
            var b = Rect;

            _form.InvokeIfRequired(x =>
            {
                x.OnBeforeDocumentPrint(theDocument, hDC, a, b);
            });
        }

        public void BeforeDocumentSave(object theDocument)
        {
            _form.BeginInvokeIfRequired(x =>
            {
                x.OnBeforeDocumentSave(theDocument);
            });
        }

        public void BeforeEnvironmentDeactivate(object theEnvironment)
        {
            _form.BeginInvokeIfRequired(x =>
            {
                x.OnBeforeEnvironmentDeactivate(theEnvironment);
            });
        }

        public void BeforeQuit()
        {
            _form.BeginInvokeIfRequired(x =>
            {
                x.OnBeforeQuit();
            });
        }

        public void BeforeWindowDeactivate(object theWindow)
        {
            _form.BeginInvokeIfRequired(x =>
            {
                x.OnBeforeWindowDeactivate(theWindow);
            });
        }

        #endregion
    }
}
