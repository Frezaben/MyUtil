package cqie.per.springtest.util.excel;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelSheet {
    String sheet();
    int sectionRow() default 0;
    int dataRow() default 1;
    boolean auto() default true;
}
