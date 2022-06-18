package com.example.msyd.assets.worker;

import lombok.AllArgsConstructor;

import java.util.List;
import java.util.Map;

/**
 * @author
 * @Description:
 * @date 2022/6/18 2:56
 */
@AllArgsConstructor
public class GenPdfAndSqlWorker extends Thread {

    private List<Map<String, String>> data;

    @Override
    public void run() {

    }

}
