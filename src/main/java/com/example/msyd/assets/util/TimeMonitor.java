package com.example.msyd.assets.util;

/**
 * @author
 * @Description: 计时器
 * @date 2022/6/14 18:34
 */
public class TimeMonitor {

    private static ThreadLocal<Long> tm = new ThreadLocal<>();

    public static void start() {
        tm.set(System.currentTimeMillis());
    }

    public static void end() {
        System.out.println("本次任务【" + Thread.currentThread().getId() + "】 共计耗时：" + (System.currentTimeMillis() - tm.get())/1000 + "s");
        tm.remove();
    }

}
