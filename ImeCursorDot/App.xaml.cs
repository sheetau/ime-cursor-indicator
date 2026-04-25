using System.Configuration;
using System.Data;
using System.Windows;

namespace ImeCursorDot
{
    public partial class App : System.Windows.Application
    {
        private static Mutex? _mutex;

        protected override void OnStartup(StartupEventArgs e)
        {
            const string name = "ImeCursorDot_SingleInstance";

            _mutex = new Mutex(true, name, out bool created);
            if (!created)
            {
                Shutdown();
                return;
            }

            base.OnStartup(e);
        }
    }
}