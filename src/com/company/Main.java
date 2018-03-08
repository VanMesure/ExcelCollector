package com.company;

import java.util.Scanner;

public class Main {

    public static void main(String[] args) {

        Searcher s = new Searcher();
        s.searchExcel();

        Excel e = new Excel();
        int status = e.collect(s);
        while(status != 0){
            System.out.println("是否尝试重新汇总上述错误文件？（y/n）：");
            Scanner scan = new Scanner(System.in);
            if(scan.next().equals("y")){
                //将出错的excel设置成s的file
                s.setExcels(e.getErrorExcels());
                e = new Excel();
                status = e.collect(s);
            }else{
                break;
            }
        }
    }
}
