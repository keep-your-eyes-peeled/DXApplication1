using DevExpress.LookAndFeel;
using DevExpress.Skins;
using DevExpress.UserSkins;
using DevExpress.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace DXApplication1
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            DevExpress.Utils.AppearanceObject.DefaultFont = new Font("Tahoma", 12, FontStyle.Regular);
            Application.Run(new Login());
            Application.ApplicationExit += delegate {
                foreach (Form f in Application.OpenForms)
                    f.Close();
            };
        }
    }
}
