package com.wxss.threadexcel.domain.entity;

import lombok.Data;

import java.util.Date;

@Data
public class Shop {
    private Long shopId;
    private String shopName;
    private Date startTime;
    private Date endTime;
}
