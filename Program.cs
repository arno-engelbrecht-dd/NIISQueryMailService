using System.ServiceProcess;
using ServiceProcess.Helpers;

namespace NIISQueryMailService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new NIISQueryMailService()
            };
            ServicesToRun.LoadServices();
        }
    }
}
