package com.example.poi;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Copyright (c) 2019 KYE Company
 * 跨越速运集团有限公司版权所有
 *
 * @author xiejiqing
 * @create 2019/7/14 0:12
 */
public class B {
    public static void main(String args[]) {
        String str = "";
        String pattern = "(|\\s{0,})";

        Pattern r = Pattern.compile(pattern);
        Matcher m = r.matcher(str);
        System.out.println(m.matches());
    }
}
