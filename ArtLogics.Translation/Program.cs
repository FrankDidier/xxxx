using ArtLogics.Translation.Properties;
using ArtLogics.Translation.View;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ArtLogics.Translation
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //System.Threading.Thread.CurrentThread.CurrentUICulture = new CultureInfo(Resources.Language);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var form = new MainForm();

            Application.Run(form);

            while (!form.Exit)
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo(form.Language);
                Thread.CurrentThread.CurrentUICulture = new CultureInfo(form.Language);
                form = new MainForm();
                Application.Run(form);
            }
        }
    }
}
