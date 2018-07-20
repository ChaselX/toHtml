package com.example.office.utils;

import org.apache.http.HttpStatus;

import java.util.HashMap;
import java.util.Map;

public class ReturnMassage extends HashMap<String, Object> {

    public  ReturnMassage(){
        put("code", 0);
        put("msg", "success");
    }
    public static ReturnMassage error() {
        return error(HttpStatus.SC_INTERNAL_SERVER_ERROR, "未知异常，请联系管理员");
    }
    public static ReturnMassage error(String msg) {
        return error(HttpStatus.SC_INTERNAL_SERVER_ERROR, msg);
    }
    public static ReturnMassage error(int code, String msg) {
        ReturnMassage r = new ReturnMassage();
        r.put("code", code);
        r.put("msg", msg);
        return r;
    }
    public static ReturnMassage ok(String msg) {
        ReturnMassage r = new ReturnMassage();
        r.put("msg", msg);
        return r;
    }

    public static ReturnMassage ok(Map<String, Object> map) {
        ReturnMassage r = new ReturnMassage();
        r.putAll(map);
        return r;
    }

    public static ReturnMassage ok() {
        return new ReturnMassage();
    }

    public ReturnMassage put(String key, Object value) {
        super.put(key, value);
        return this;
    }
}
