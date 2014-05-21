using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace Example1
{
    public static class ControlExtensions
    {
        /// <summary>
        /// Synchornous call to control but may load to deadlock if events are firing rapidly.
        /// </summary>
        public static void InvokeIfRequired<TControl>(this TControl control, Action<TControl> action)
            where TControl : Control
        {
            if (control.InvokeRequired)
            {
                control.Invoke(action, control);
            }
            else
            {
                action(control);
            }
        }

        /// <summary>
        /// Asynchornous to control and should be safest.
        /// </summary>
        public static void BeginInvokeIfRequired<TControl>(this TControl control, Action<TControl> action)
            where TControl : Control
        {
            if (control.InvokeRequired)
            {
                control.BeginInvoke(action, control);
            }
            else
            {
                action(control);
            }
        }
    }
}
