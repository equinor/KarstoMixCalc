using System;
using System.Globalization;
using Topshelf;

namespace MixCalc
{
    class Program
    {
        static void Main()
        {
            var exitCode = HostFactory.Run(x =>
            {
                x.Service<MixCalcService>(s =>
                {
                    s.ConstructUsing(MixCalcService => new MixCalcService());
                    s.WhenStarted(MixCalcService => MixCalcService.Start());
                    s.WhenStopped(MixCalcService => MixCalcService.Stop());
                });

                x.SetServiceName("MixCalcService");
            });

            int exitCodeValue = (int)Convert.ChangeType(exitCode, exitCode.GetTypeCode(), CultureInfo.InvariantCulture);
            Environment.ExitCode = exitCodeValue;
        }
    }
}
