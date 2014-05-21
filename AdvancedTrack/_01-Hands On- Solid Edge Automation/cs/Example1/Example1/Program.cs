using SolidEdgeContrib;
using SolidEdgeContrib.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Example1
{
    class Program
    {
        [STAThread] // <-- In a console application, you must add the STAThread attribute!
        static void Main(string[] args)
        {

            // Note that we're using SolidEdgeContrib.OleMessageFilter.
            OleMessageFilter.Register();

            // On a system where Solid Edge is installed, the COM ProgID will be
            // defined in registry: HKEY_CLASSES_ROOT\SolidEdge.Application
            Type t = Type.GetTypeFromProgID(SolidEdge.PROGID.Application, throwOnError: true);

            // Using the discovered Type, create and return a new instance of Solid Edge.
            var application = (SolidEdgeFramework.Application)Activator.CreateInstance(type: t);

            // Show Solid Edge.
            application.Visible = true;

            // Get a reference to the Documents collection.
            var documents = application.Documents;

            // Demonstrate SolidEdgeContrib provided method extensions.
            DemoApplicationExtensionMethods(application, documents);

            // Turn off Solid Edge prompts.
            application.DisplayAlerts = false;

            // Close all documents without saving.
            documents.Close();

            // Terminate Solid Edge.
            application.Quit();

            // Note that we're using SolidEdgeContrib.OleMessageFilter.
            OleMessageFilter.Unregister();
        }

        static void DemoApplicationExtensionMethods(SolidEdgeFramework.Application application, SolidEdgeFramework.Documents documents)
        {
            // Note the extension methods are only available when you use:
            // using SolidEdgeContrib.Extensions;

            // Add an assembly document.
            var assemblyDocument = documents.AddAssemblyDocument();

            // Always good to call DoIdle() after creating a new document.
            application.DoIdle();

            // Get a SolidEdgeFramework.SolidEdgeDocument typed reference to the active document.
            var activeDocument = application.GetActiveDocument();

            // Demonstrate generic GetActiveDocument extension method.
            var activeAssemblyDocument = application.GetActiveDocument<SolidEdgeAssembly.AssemblyDocument>();

            // Always good to call DoIdle() after creating a new document.
            application.DoIdle();

            // Add a draft document.
            var draftDocument = documents.AddDraftDocument();

            // Always good to call DoIdle() after creating a new document.
            application.DoIdle();

            // Demonstrate generic GetActiveDocument extension method.
            var activeDraftDocument = application.GetActiveDocument<SolidEdgeDraft.DraftDocument>();

            // Add a part document.
            var partDocument = documents.AddPartDocument();

            // Always good to call DoIdle() after creating a new document.
            application.DoIdle();

            // Demonstrate generic GetActiveDocument extension method.
            var activePartDocument = application.GetActiveDocument<SolidEdgePart.PartDocument>();

            // Get a reference to the RefPlanes collection.
            var refPlanes = activePartDocument.RefPlanes;

            // Demonstrate using extension methods to easily get specific RefPlanes.
            var frontPlane = refPlanes.GetFrontPlane();
            var rightPlane = refPlanes.GetRightPlane();
            var topPlane = refPlanes.GetTopPlane();

            // Add a sheet metal document.
            var sheetMetalDocument = documents.AddSheetMetalDocument();

            // Always good to call DoIdle() after creating a new document.
            application.DoIdle();

            // Demonstrate generic GetActiveDocument extension method.
            var activeSheetMetalDocument = application.GetActiveDocument<SolidEdgePart.SheetMetalDocument>();

            // Demonstrate StartCommand extension methods.
            application.StartCommand(SolidEdgeConstants.SheetMetalCommandConstants.SheetMetalViewBackView);
            application.StartCommand(SolidEdgeConstants.SheetMetalCommandConstants.SheetMetalViewBottomView);
            application.StartCommand(SolidEdgeConstants.SheetMetalCommandConstants.SheetMetalViewClipping);
            application.StartCommand(SolidEdgeConstants.SheetMetalCommandConstants.SheetMetalViewDimetricView);
            application.StartCommand(SolidEdgeConstants.SheetMetalCommandConstants.SheetMetalViewFrontView);
            application.StartCommand(SolidEdgeConstants.SheetMetalCommandConstants.SheetMetalViewISOView);
            application.StartCommand(SolidEdgeConstants.SheetMetalCommandConstants.SheetMetalViewLeftView);

            // Get a strongly typed SolidEdgeFramework.ISEApplicationEvents_Event.
            var applicationEvents = application.GetApplicationEvents();

            // Get a strongly typed reference to the active environment.
            var environment = application.GetActiveEnvironment();

            // Get the seApplicationGlobalSystemInfo global parameter. Return type is an object.
            var globalSystemInfoObject = application.GetGlobalParameter(SolidEdgeFramework.ApplicationGlobalConstants.seApplicationGlobalSystemInfo);

            // Get the seApplicationGlobalSystemInfo global parameter. Using the generic overload, return type is a string.
            var globalSystemInfoString = application.GetGlobalParameter<string>(SolidEdgeFramework.ApplicationGlobalConstants.seApplicationGlobalSystemInfo);

            // Get an instance of a System.Diagnostics.Process that represents the Edge.exe process.
            var process = application.GetProcess();

            // Get an instance of a System.Version that represents the version of Solid Edge rather than just a string.
            var version = application.GetVersion();

        }
    }
}
