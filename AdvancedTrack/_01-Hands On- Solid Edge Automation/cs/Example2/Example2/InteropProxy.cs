using SolidEdgeContrib;
using SolidEdgeContrib.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting;
using System.Text;

namespace Example2
{
    public class InteropProxy : MarshalByRefObject
    {
        /// <summary>
        /// Executes interop code in an isolated AppDomain.
        /// </summary>
        /// <remarks>
        /// Notice that we do not have to worry about RCW's. i.e. Marshal.ReleaseComObject.
        /// </remarks>
        public void DoIsolatedTask(SolidEdgeFramework.Application applicationTransparentProxy)
        {
            // See what AppDomain we're currently executing in.
            var currentAppDomain = AppDomain.CurrentDomain;

            // This will never be the default AppDomain at this point.
            var isDefaultAppDomain = currentAppDomain.IsDefaultAppDomain();

            // Register with OLE to handle concurrency issues on the current thread.
            OleMessageFilter.Register();

            // RCW's cross AppDomains as TransparentProxies.
            // Unwrap the TransparentProxy object that was passed into this AppDomain.
            var application = UnwrapTransparentProxy<SolidEdgeFramework.Application>(applicationTransparentProxy);
            var documents = application.Documents;

            // Add a new part document.
            var partDocument = documents.AddPartDocument();

            // Always a good idea to give SE a chance to breathe.
            application.DoIdle();

            // Optional performance improvement tweaks.
            application.DelayCompute = true;
            application.ScreenUpdating = false;

            // Create a polygon in the part document.
            CreatePolygon(partDocument);

            // Undo performance improvement tweaks.
            application.DelayCompute = false;
            application.ScreenUpdating = true;

            // Register with OLE to handle concurrency issues on the current thread.
            OleMessageFilter.Unregister();
        }

        private void CreatePolygon(SolidEdgePart.PartDocument partDocument)
        {
            // Get a reference to the Application object.
            var application = partDocument.Application;

            // Get a reference to the RefPlanes collection.
            var refPlanes = partDocument.RefPlanes;

            // Get a reference to the top RefPlane using extension method.
            var refPlane = refPlanes.GetTopPlane();

            // Get a reference to the ProfileSets collection.
            var profileSets = partDocument.ProfileSets;

            // Add a new ProfileSet.
            var profileSet = profileSets.Add();

            // Get a reference to the Profiles collection.
            var profiles = profileSet.Profiles;

            // Add a new Profile.
            var profile = profiles.Add(refPlane);

            // Get a reference to the Relations2d collection.
            var relations2d = (SolidEdgeFrameworkSupport.Relations2d)profile.Relations2d;

            // Get a reference to the Lines2d collection.
            var lines2d = profile.Lines2d;

            int sides = 8;
            double angle = 360 / sides;
            angle = (angle * Math.PI) / 180;

            double radius = .05;
            double lineLength = 2 * radius * (Math.Tan(angle) / 2);

            // x1, y1, x2, y2
            double[] points = { 0.0, 0.0, 0.0, 0.0 };

            double x = 0.0;
            double y = 0.0;

            points[2] = -((Math.Cos(angle / 2) * radius) - x);
            points[3] = -((lineLength / 2) - y);

            // Draw each line.
            for (int i = 0; i < sides; i++)
            {
                points[0] = points[2];
                points[1] = points[3];
                points[2] = points[0] + (Math.Sin(angle * i) * lineLength);
                points[3] = points[1] + (Math.Cos(angle * i) * lineLength);

                lines2d.AddBy2Points(points[0], points[1], points[2], points[3]);
            }

            // Create endpoint relationships.
            for (int i = 1; i <= lines2d.Count; i++)
            {
                if (i == lines2d.Count)
                {
                    relations2d.AddKeypoint(lines2d.Item(i), (int)SolidEdgeConstants.KeypointIndexConstants.igLineEnd, lines2d.Item(1), (int)SolidEdgeConstants.KeypointIndexConstants.igLineStart);
                }
                else
                {
                    relations2d.AddKeypoint(lines2d.Item(i), (int)SolidEdgeConstants.KeypointIndexConstants.igLineEnd, lines2d.Item(i + 1), (int)SolidEdgeConstants.KeypointIndexConstants.igLineStart);
                    relations2d.AddEqual(lines2d.Item(i), lines2d.Item(i + 1));
                }
            }

            // Get a reference to the ActiveSelectSet.
            var selectSet = application.ActiveSelectSet;

            // Empty ActiveSelectSet.
            selectSet.RemoveAll();
            
            // Add all lines to ActiveSelectSet.
            for (int i = 1; i <= lines2d.Count; i++)
            {
                selectSet.Add(lines2d.Item(i));
            }

            // Switch to ISO view.
            application.StartCommand(SolidEdgeConstants.PartCommandConstants.PartViewISOView);
        }

        private T UnwrapTransparentProxy<T>(object rcw) where T : class
        {
            if (RemotingServices.IsTransparentProxy(rcw))
            {
                IntPtr punk = Marshal.GetIUnknownForObject(rcw);

                try
                {
                    return (T)Marshal.GetObjectForIUnknown(punk);
                }
                finally
                {
                    Marshal.Release(punk);
                }
            }

            return (T)rcw;
        }
    }
}
