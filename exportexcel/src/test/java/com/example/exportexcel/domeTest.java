/**
 * Copyright (C), 2015-2019, XXX有限公司
 * FileName: domeTest
 * Author:   jcf
 * Date:     2019/6/2 14:24
 * Description:
 * History:
 * <author>          <time>          <version>          <desc>
 * 作者姓名           修改时间           版本号              描述
 */

package com.example.exportexcel;


import java.util.ArrayList;
import java.util.List;

public class domeTest {
    public static void main(String[] args) {
        List<String> xStrs = new ArrayList<>();
        List<String> yStrs = new ArrayList<>();
        int year0 = 2012;
        int month = 1;
        for (int i = 0; i < 60; i++) {
            if(month>12){
                month = 1;
                year0 ++;
            }
            String ym = String.valueOf(year0);
            if(month<10){
                ym = ym + "0"+String.valueOf(month);
            }else {
                ym = ym+String.valueOf(month);
            }
            month++;
            xStrs.add(ym);
        }

        for (String str : xStrs) {
            System.out.print(str + "   ");
        }
    }
}
