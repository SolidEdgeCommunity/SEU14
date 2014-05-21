using SolidEdgeContrib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Example1
{
    public partial class Form1 : Form
    {
        private SolidEdgeFramework.Application _application = null;
        private ApplicationEventWatcher _applicationEventWatcher = null;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            OleMessageFilter.Register();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_applicationEventWatcher != null)
            {
                _applicationEventWatcher.Dispose();
                _applicationEventWatcher = null;
            }
            _application = null;
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void eventButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (eventButton.Checked)
                {
                    if (_application == null)
                    {
                        // On a system where Solid Edge is installed, the COM ProgID will be
                        // defined in registry: HKEY_CLASSES_ROOT\SolidEdge.Application
                        Type t = Type.GetTypeFromProgID(SolidEdge.PROGID.Application, throwOnError: true);

                        // Using the discovered Type, create and return a new instance of Solid Edge.
                        _application = (SolidEdgeFramework.Application)Activator.CreateInstance(type: t);

                        // Show Solid Edge.
                        _application.Visible = true;
                    }

                    _applicationEventWatcher = new ApplicationEventWatcher(this, _application);
                }
                else
                {
                    _applicationEventWatcher.Dispose();
                    _applicationEventWatcher = null;
                    _application = null;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void clearButton_Click(object sender, EventArgs e)
        {
            eventLogTextBox.Clear();
        }

        #region SolidEdgeFramework.ISEApplicationEvents

        public void OnAfterActiveDocumentChange(object theDocument)
        {
            eventLogTextBox.AppendText("OnAfterActiveDocumentChange");
            eventLogTextBox.AppendText(Environment.NewLine);
        }

        public void OnAfterCommandRun(int theCommandID)
        {
            eventLogTextBox.AppendText("OnAfterCommandRun");
            eventLogTextBox.AppendText(Environment.NewLine);
        }

        public void OnAfterDocumentOpen(object theDocument)
        {
            eventLogTextBox.AppendText("OnAfterDocumentOpen");
            eventLogTextBox.AppendText(Environment.NewLine);
        }

        public void OnAfterDocumentPrint(object theDocument, int hDC, double ModelToDC, int Rect)
        {
            eventLogTextBox.AppendText("OnAfterDocumentPrint");
            eventLogTextBox.AppendText(Environment.NewLine);
        }

        public void OnAfterDocumentSave(object theDocument)
        {
            eventLogTextBox.AppendText("OnAfterDocumentSave");
            eventLogTextBox.AppendText(Environment.NewLine);
        }

        public void OnAfterEnvironmentActivate(object theEnvironment)
        {
            eventLogTextBox.AppendText("OnAfterEnvironmentActivate");
            eventLogTextBox.AppendText(Environment.NewLine);
        }

        public void OnAfterNewDocumentOpen(object theDocument)
        {
            eventLogTextBox.AppendText("OnAfterNewDocumentOpen");
            eventLogTextBox.AppendText(Environment.NewLine);
        }

        public void OnAfterNewWindow(object theWindow)
        {
            eventLogTextBox.AppendText("OnAfterNewWindow");
            eventLogTextBox.AppendText(Environment.NewLine);
        }

        public void OnAfterWindowActivate(object theWindow)
        {
            eventLogTextBox.AppendText("OnAfterWindowActivate");
            eventLogTextBox.AppendText(Environment.NewLine);
        }

        public void OnBeforeCommandRun(int theCommandID)
        {
            eventLogTextBox.AppendText("OnBeforeCommandRun");
            eventLogTextBox.AppendText(Environment.NewLine);
        }

        public void OnBeforeDocumentClose(object theDocument)
        {
            eventLogTextBox.AppendText("OnBeforeDocumentClose");
            eventLogTextBox.AppendText(Environment.NewLine);
        }

        public void OnBeforeDocumentPrint(object theDocument, int hDC, double ModelToDC, int Rect)
        {
            eventLogTextBox.AppendText("OnBeforeDocumentPrint");
            eventLogTextBox.AppendText(Environment.NewLine);
        }

        public void OnBeforeDocumentSave(object theDocument)
        {
            eventLogTextBox.AppendText("OnBeforeDocumentSave");
            eventLogTextBox.AppendText(Environment.NewLine);
        }

        public void OnBeforeEnvironmentDeactivate(object theEnvironment)
        {
            eventLogTextBox.AppendText("OnBeforeEnvironmentDeactivate");
            eventLogTextBox.AppendText(Environment.NewLine);
        }

        public void OnBeforeQuit()
        {
            eventLogTextBox.AppendText("OnBeforeQuit");
            eventLogTextBox.AppendText(Environment.NewLine);

            _applicationEventWatcher.Dispose();
            _applicationEventWatcher = null;
            _application = null;
        }

        public void OnBeforeWindowDeactivate(object theWindow)
        {
            eventLogTextBox.AppendText("OnBeforeWindowDeactivate");
            eventLogTextBox.AppendText(Environment.NewLine);
        }

        #endregion

    }
}
