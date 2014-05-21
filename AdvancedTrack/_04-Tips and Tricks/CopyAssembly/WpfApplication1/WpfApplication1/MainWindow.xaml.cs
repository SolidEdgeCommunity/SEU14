using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using SolidEdge.Framework;
using SolidEdge.FrameworkSupport;
using SolidEdge.Assembly;
using System.IO;

namespace WpfApplication1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public const string SolidEdgeApplication = "SolidEdge.Application";

        private SolidEdge.Framework.Interop.Application oSolidEdge = null;

        public MainWindow()
        {
            InitializeComponent();
            TargetFolder.Text = Properties.Settings.Default["CopyPath"].ToString();
            AttachToSolidEdge();
        }

        private void buttonCopy_Click(object sender, RoutedEventArgs e)
        {
            DirectoryInfo di = new DirectoryInfo(TargetFolder.Text.ToString());
            if (!di.Exists)
            {
                System.Windows.Forms.MessageBox.Show(TargetFolder.Text.ToString() + " does not exist");
                return;
            }

            if (FileNames.Items.IsEmpty)
                AttachToSolidEdge();

            int n = 0;
            string[] copies = new string[FileNames.Items.Count];

            foreach (string source in FileNames.Items)
            {
                string target = System.IO.Path.Combine(TargetFolder.Text, System.IO.Path.GetFileName(source));
                try
                {
                    System.IO.File.Copy(source, target, true);
                    copies[n++] = target;
                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show("Error copying " + source + " to " + target + ": " + ex.ToString());
                }
            }

            FileNames.Items.Clear();
            foreach (string target in copies)
            {
                FileNames.Items.Add(target);
            }

            Properties.Settings.Default["CopyPath"] = TargetFolder.Text;
            Properties.Settings.Default.Save();
        }

        private void buttonBrowse_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog browse = new System.Windows.Forms.FolderBrowserDialog();

            string defpath = Properties.Settings.Default["CopyPath"].ToString();
            DirectoryInfo di = new DirectoryInfo(defpath);
            if (di.Exists)
                browse.SelectedPath = defpath;
            
            if (System.Windows.Forms.DialogResult.OK == browse.ShowDialog())
            {
                TargetFolder.Text = browse.SelectedPath;
                Properties.Settings.Default["CopyPath"] = TargetFolder.Text;
                Properties.Settings.Default.Save();
            }
        }

        private void buttonAttach_Click(object sender, RoutedEventArgs e)
        {
            AttachToSolidEdge();
        }

        private void AttachToSolidEdge()
        {
            FileNames.Items.Clear();

            try
            {
                object oApp = Marshal.GetActiveObject(SolidEdgeApplication);
                oSolidEdge = (SolidEdge.Framework.Interop.Application)oApp;
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Failed to connect to Solid Edge: " + ex.ToString());
                return;
            }

            int nDocs = oSolidEdge.Documents.Count;

            try
            {
                if (oSolidEdge.ActiveDocumentType == SolidEdge.Framework.Interop.DocumentTypeConstants.igAssemblyDocument)
                {
                    object oDoc = oSolidEdge.ActiveDocument;
                    SolidEdge.Assembly.Interop.AssemblyDocument asmDoc = (SolidEdge.Assembly.Interop.AssemblyDocument)oDoc;
                    string sFullName = asmDoc.FullName;
                    if (!FileNames.Items.Contains(sFullName))
                        FileNames.Items.Add(sFullName);
                    EnumerateDocuments(asmDoc);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Failed to enumerate documents: " + ex.ToString());
            }
        }
        
        private void EnumerateDocuments(SolidEdge.Assembly.Interop.AssemblyDocument asmDoc)
        {
            int nOccs = asmDoc.Occurrences.Count;
            foreach (SolidEdge.Assembly.Interop.Occurrence occ in asmDoc.Occurrences)
            {
                if ((occ.Visible == false) && (CopyHiddenFiles.IsChecked == false))
                    continue;

                string sFullName = occ.OccurrenceFileName;
                if (!FileNames.Items.Contains(sFullName))
                    FileNames.Items.Add(sFullName);
                object oSourceDoc = occ.OccurrenceDocument;
                if (oSourceDoc != null)
                {
                    try
                    {
                        SolidEdge.Assembly.Interop.AssemblyDocument subAsmDoc = (SolidEdge.Assembly.Interop.AssemblyDocument)oSourceDoc;
                        EnumerateDocuments(subAsmDoc);
                    }
                    catch (System.Exception /*ex*/)
                    {
                    	
                    }
                }
            }
        }

        private void CopyHiddenFiles_Checked(object sender, RoutedEventArgs e)
        {
            AttachToSolidEdge();
        }
    }
}
