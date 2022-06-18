package com.example.msyd.assets.formatter;

import lombok.Data;

import java.text.SimpleDateFormat;

/**
 * @author
 * @Description: 日期格式化
 * @date 2022/6/13 18:18
 */
@Data
//@AllArgsConstructor
public class DateColumnFormatter implements ColumnFormatter<String> {

//    private Object target;

    private ThreadLocal<SimpleDateFormat> df = new ThreadLocal<SimpleDateFormat>() {
        @Override
        protected SimpleDateFormat initialValue() {
            return new SimpleDateFormat("yyyy-MM-dd");
        }
    };

    @Override
    public String format(Object target) {
        return df.get().format(target);
    }
}
