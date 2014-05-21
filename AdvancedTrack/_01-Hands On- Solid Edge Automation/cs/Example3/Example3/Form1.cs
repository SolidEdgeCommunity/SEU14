using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace Example3
{
    public partial class Form1 : Form
    {
        private SolidEdgeFramework.Application _application;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Disable the button.
            button1.Enabled = false;

            // Toggle the label visible state.
            label1.Visible = !label1.Visible;

            // Get a reference to Solid Edge if we don't already have one.
            if (_application == null)
            {
                try
                {
                    // Attempt to connect to a running instace.
                    _application = (SolidEdgeFramework.Application)Marshal.GetActiveObject(SolidEdge.PROGID.Application);
                }
                catch
                {
                }
            }

            // See what AppDomain we're currently executing in.
            var currentAppDomain = AppDomain.CurrentDomain;

            // This will always be the default AppDomain at this point.
            var isDefaultAppDomain = currentAppDomain.IsDefaultAppDomain();

            backgroundWorker1.RunWorkerAsync(_application);
        }

        private void CreateSeparateAppDomainAndExecuteIsolatedTask(SolidEdgeFramework.Application application)
        {
            AppDomain interopDomain = null;

            try
            {
                var thread = new System.Threading.Thread(() =>
                {
                    // Create a custom AppDomain to do COM Interop.
                    interopDomain = AppDomain.CreateDomain("Interop Domain");

                    Type proxyType = typeof(InteropProxy);

                    // Create a new instance of InteropProxy in the isolated application domain.
                    InteropProxy interopProxy = interopDomain.CreateInstanceAndUnwrap(
                        proxyType.Assembly.FullName,
                        proxyType.FullName) as InteropProxy;

                    try
                    {
                        interopProxy.DoIsolatedTask(application);
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.StackTrace, ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                });

                // Important! Set thread apartment state to STA.
                thread.SetApartmentState(System.Threading.ApartmentState.STA);

                // Start the thread.
                thread.Start();

                // Wait for the thead to finish.
                thread.Join();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.StackTrace, ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (interopDomain != null)
                {
                    // Unload the Interop AppDomain. This will automatically free up any COM references.
                    AppDomain.Unload(interopDomain);
                }
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            var application = (SolidEdgeFramework.Application)e.Argument;

            // Check to see if we were passed a valid application instance.
            if (application == null)
            {
                // On a system where Solid Edge is installed, the COM ProgID will be
                // defined in registry: HKEY_CLASSES_ROOT\SolidEdge.Application
                Type t = Type.GetTypeFromProgID(SolidEdge.PROGID.Application, throwOnError: true);

                // Using the discovered Type, create and return a new instance of Solid Edge.
                application = (SolidEdgeFramework.Application)Activator.CreateInstance(type: t);
            }

            // Make sure Solid Edge is visible.
            application.Visible = true;

            // Create a separate AppDomain and execute our code.
            CreateSeparateAppDomainAndExecuteIsolatedTask(application);
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // Hide the label.
            label1.Visible = false;

            // Enable the button.
            button1.Enabled = true;
        }
    }
}
