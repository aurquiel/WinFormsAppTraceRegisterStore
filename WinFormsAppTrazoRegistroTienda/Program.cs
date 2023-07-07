namespace WinFormsAppTrazoRegistroTienda
{
    internal static class Program
    {
        // Mutex can be made static so that GC doesn't recycle
        // same effect with GC.KeepAlive(mutex) at the end of main
        static Mutex mutex = new Mutex(false, "Aplicacion-Trazo-Registros");
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // if you like to wait a few seconds in case that the instance is just 
            // shutting down
            if (!mutex.WaitOne(TimeSpan.FromSeconds(2), false))
            {
                MessageBox.Show("Ya la aplicacion esta corriendo.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // To customize application configuration such as set high DPI settings or default font,
                // see https://aka.ms/applicationconfiguration.
                ApplicationConfiguration.Initialize();
                Application.Run(new LoginForm());
            }
            finally { mutex.ReleaseMutex(); } // I find this more explicit

            
        }
    }
}