package cqie.per.springtest.util.excel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCell {
    String cell();
    boolean useIndex() default false;
    boolean notNull() default false;
    boolean readIgnore() default false;
    boolean outputIgnore() default false;
    boolean mergedSameRegion() default false;
}