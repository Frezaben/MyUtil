package cqie.per.springtest.entity;

import lombok.Data;

import java.math.BigDecimal;
import java.util.List;
import java.util.Map;
import java.util.Set;

@Data
public class JsonTest {
    private String Str;
    private Integer numInt;
    private Long numLon;
    private BigDecimal numBigD;
    private Integer[] nums;
    private List<String> listString;
    private Set<Long> setLong;
    private Map<String,Integer> map1;
    private Map<Integer,String> map2;
}
