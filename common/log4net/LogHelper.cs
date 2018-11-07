//需要在 AssemblyInfo.cs文件配置
//[assembly: log4net.Config.XmlConfigurator(ConfigFile = "log4net.config", Watch = true)]
public static class LogHelper
    {
        static log4net.ILog log = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        public static void Err(object message, string exception)
        {
            log.Error(message, new Exception(exception));
        }

        public static void Info(object message, string exception)
        {
            log.Info(message, new Exception(exception));
        }
    }