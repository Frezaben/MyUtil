package cqie.per.springtest.entity;

import cqie.per.springtest.util.excel.annotations.ExcelCell;
import cqie.per.springtest.util.excel.annotations.ExcelSheet;
import cqie.per.springtest.util.excel.enums.ReadModelEnum;

@ExcelSheet(sheet = "名称",model = ReadModelEnum.MANUAL)
public class StudentInfo {
    @ExcelCell(cell = "A",mergedSameRegion = true, outputNotNull = true)
    private String name;
    @ExcelCell(cell = "B",mergedSameRegion = true)
    private Integer uid;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getUid() {
        return uid;
    }

    public void setUid(Integer uid) {
        this.uid = uid;
    }

    @Override
    public String toString() {
        return "StudentInfo{" +
                "name='" + name + '\'' +
                ", uid=" + uid +
                '}';
    }
}
