package com.retheviper.excel.definition.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Mark class as definition of Sheet.
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface Sheet {

    /**
     * Sheet's name. If null or "", it will be considered as class's name.
     */
    String name() default "";

    /**
     * Sheet's index.
     */
    int index() default -1;

    /**
     * Row index which data part begins.
     */
    int dataStartIndex() default 0;
}
