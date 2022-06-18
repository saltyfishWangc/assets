package com.example.msyd.assets.formatter;

/**
 * @author
 * @Description: 列格式化接口定义
 * @date 2022/6/13 18:16
 */
public interface ColumnFormatter<T> {

    T format(Object obj);
}
