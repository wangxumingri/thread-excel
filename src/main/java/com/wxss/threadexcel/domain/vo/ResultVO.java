package com.wxss.threadexcel.domain.vo;

import lombok.Data;

@Data
public class ResultVO<T> {
    private Boolean success;
    private Integer code;
    private String msg;
    private T data;

    public static ResultVO<Object> success() {
        return ResultVO.success(null);
    }

    public static <T> ResultVO<T> success(T data) {
        ResultVO<T>  resultVO = new ResultVO<>();
        resultVO.setSuccess(Boolean.TRUE);
        resultVO.setCode(100000);
        resultVO.setData(data);

        return resultVO;
    }

    public static <T> ResultVO<T> fail() {
        return ResultVO.fail(null);
    }
    public static <T> ResultVO<T> fail(String msg) {
        ResultVO<T>  resultVO = new ResultVO<>();
        resultVO.setSuccess(Boolean.FALSE);
        resultVO.setCode(999999);
        resultVO.setMsg(msg);
        return resultVO;
    }
}
