package cqie.per.springtest.entity;

import cqie.per.springtest.util.excel.annotations.ExcelCell;
import cqie.per.springtest.util.excel.annotations.ExcelSheet;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;

import java.util.Date;
import java.util.List;
import java.util.Set;

@ExcelSheet(sheet = "工作表1")
@Data
public class UserInfo {
    @ApiModelProperty("姓名")
    private String name;
    private List<String> list;
    private Set<Integer> set;
    @ApiModelProperty("编号")
    private Long id;
    @ExcelCell(cell = "用户代码")
    private Integer uid;
    public String age;
    @ApiModelProperty("时间")
    public Date time;
    @ApiModelProperty("日期")
    public String date;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public Integer getUid() {
        return uid;
    }

    public void setUid(Integer uid) {
        this.uid = uid;
    }

    public String getAge() {
        return age;
    }

    public void setAge(String age) {
        this.age = age;
    }
}
