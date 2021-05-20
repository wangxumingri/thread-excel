package com.wxss.threadexcel.utils;

import java.util.concurrent.LinkedBlockingDeque;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;

public class MyThreadPool {
    private static class Holder{
        static ThreadPoolExecutor threadPoolExecutor = new ThreadPoolExecutor(6, 10, 1000L, TimeUnit.SECONDS, new LinkedBlockingDeque<>());
    }
    public static ThreadPoolExecutor getThreadPool(){
        return Holder.threadPoolExecutor;
    }
}
