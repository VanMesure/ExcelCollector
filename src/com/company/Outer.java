package com.company;

import java.io.File;
import java.util.Iterator;
import java.util.List;

public class Outer {
    public void outFileInfo(File module, List<File> excels){
        if(module == null){
            System.out.println("模板文件缺失！（请将模板文件命名为 module.xlsx）");
            System.exit(0);
            return;
        }
        if(excels.size() == 0){
            System.out.println("未找到任何非模板的xlsx文件！ 请将所有文件放到当前目录下");
            System.exit(0);
            return;
        }
        System.out.println("找到如下文件：");
        System.out.println("---------------");
        for(File f : excels){
            System.out.println(f.getName());
        }
        System.out.println("--------------");
        System.out.println("正在将上述文件汇总到module.xlsx .....");
        System.out.println("");
    }

    /*
    * 如果当前目录下全部excel汇总完毕，则返回0，否则返回失败的文件数*/
    public int outCollectResult(int success, List<File> failList){
        System.out.println("\n汇总完毕，成功汇总" + success + "个文件，剩余" + failList.size() + "个文件未成功");
        if(failList.size() != 0){
            System.out.println("以下是未被成功汇总的文件，请检查格式：");
            for(File f : failList){
                System.out.println("--" + f.getName());
            }
            return failList.size();
        }
        return 0;
    }

}
