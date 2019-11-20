package oneplay.office4j.utils;

import org.apache.logging.log4j.Level;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.core.LoggerContext;
import org.apache.logging.log4j.core.config.Configuration;
import org.apache.logging.log4j.core.config.LoggerConfig;


public class Log4j2Util {

    public static void setLoggerLevel(Class<?> clazz, String level) {
        String loggerName = clazz.getName();
        LoggerContext loggerContext = (LoggerContext) LogManager.getContext(false);
        Configuration configuration = loggerContext.getConfiguration();
        LoggerConfig loggerConfig = configuration.getLoggerConfig(loggerName);
        loggerConfig.setLevel(Level.valueOf(level.toUpperCase().replaceAll("ING", "")));
    }

}
