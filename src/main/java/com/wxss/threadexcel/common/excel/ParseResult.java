package com.wxss.threadexcel.common.excel;

import lombok.Data;

import java.util.List;

public class ParseResult<R> {
    private List<R> data;
    private Boolean success;
    private String errMsg;



    public List<R> getData() {
        return data;
    }

    public void setData(List<R> data) {
        this.data = data;
    }

    public Boolean getSuccess() {
        return success;
    }

    public void setSuccess(Boolean success) {
        this.success = success;
    }

    public String getErrMsg() {
        return errMsg;
    }

    public void setErrMsg(String errMsg) {
        this.errMsg = errMsg;
    }


    @Override
    public String toString() {
        return "ParseResult{" +
                "data=" + data +
                ", success=" + success +
                ", errMsg='" + errMsg + '\'' +
                '}';
    }
}
