package com.company;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class Searcher {
    private String nowPath;                                 //当前路径
    private List<File> excels = new ArrayList<File>();      //需要汇总的文件
    private File module;                                    //模板文件
    private Outer out = new Outer();
    Searcher(){
        nowPath = System.getProperty("user.dir");
    }

    public void searchExcel(){
        File dir = new File(nowPath + "\\excels");
        File[] files = dir.listFiles();
        List<File> excels = new ArrayList<>();
        //判断一下目录是否存在
        if(!dir.exists()){
            System.out.println("未检测到excels文件夹！已自动创建，请将excel文件放入excels文件夹中并重新启动程序");
            dir.mkdir();
            System.exit(1);
        }

        //遍历当前目录下的文件，找出模板文件和普通文件
        for(File f: files){
            String fileName = f.getName();
            if(fileName.equals("module.xlsx")){
                this.module = f;
            }else if(fileName.substring(fileName.lastIndexOf(".")).equals(".xlsx")){
                excels.add(f);
            }
        }
        this.excels = excels;

        //输出信息
        out.outFileInfo(module, this.excels);

    }

    public File getModule(){
        return this.module;
    }
    public List<File> getExcels(){
        return  this.excels;
    }

    public void setExcels(List<File> excels) {
        this.excels = excels;
    }

}
