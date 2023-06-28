package com.example.config;

import org.springframework.context.ApplicationContext;

public class SpringBeanUtil {
    public static ApplicationContext applicationContext;
    /**
     * 通过class获取Bean.
     *
     * @param clazz
     * @param <T>
     * @return
     */
    public static <T> T getBean(Class<T> clazz) {
        return applicationContext.getBean(clazz);
    }
}
