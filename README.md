# qc-excel

#### 介绍
easyExcel常用导出导出工具类，可以简化操作

#### 软件架构
springboot+easy excel


#### 安装教程

1.  git clone xxx
2.  切换到dev分支

#### 使用说明

1. maven打包至本地仓库或其他私有仓库
2. 导入qcexcel-spring-boot-starter依赖

代码1：使用类
~~~ java
@Override
    public void importExcel(MultipartFile file) {
        try {
            Map<Integer,String > headMap = new ConcurrentHashMap<>();
            headMap.put(0,"账号");
            headMap.put(1,"用户名");
            // importExcel见代码2
            EasyExcelUtil.importExcel(file.getInputStream(),importExcel.getClass(),"importFile", UserInfo.class,null,headMap);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Override
    public void exportFile(HttpServletResponse response) {
        String fileName="/excel/导出模板.xlsx";
        Map<String,Object> importContent= new HashMap<>();
        Map<String,Object> map=new HashMap<>();
        map.put("title","测试");
        map.put("soOrderCode","123344");
        map.put("createPersonName","管理员");
        map.put("mobilPhone","19922220000");
        map.put("requiredCompletionTime",new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date()));
        map.put("netDispatchType","设备类资源预占");
        map.put("major","接入网");
        map.put("isCheck","是");
        map.put("businessOpenType","测试");
        List<Map<String,Object>> list = new ArrayList<>();
        list.add(new HashMap<String,Object>(){{
            put("baseStationZ","123444");
            put("deviceTypeA","类型");
            put("deviceTypeZ","Z类型");
        }});
        importContent.put("normal",map);
        importContent.put("accept",list);
        try {
            BufferedOutputStream outputStream = new BufferedOutputStream(new FileOutputStream("H:/test/导出测试.xlsx"));
            EasyExcelUtil.complexExport(fileName,"导出测试","调度",outputStream,importContent);
            EasyExcelUtil.complexExport(fileName,"导出测试","调度",response,importContent);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    public void exportMoreFile(HttpServletResponse response) {
        String fileName="/excel/导出模板.xlsx";
        Map<String,Object> importContent1= new HashMap<>();
        Map<String,Object> importContent2= new HashMap<>();
        Map<String,Object> map=new HashMap<>();
        List<Map<String,Object>> list1 = new ArrayList<>();
        List<Map<String,Object>> list2 = new ArrayList<>();
        map.put("title","测试");
        map.put("soOrderCode","123344");
        map.put("createPersonName","管理员");
        map.put("mobilPhone","19922220000");
        map.put("requiredCompletionTime",new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date()));
        map.put("netDispatchType","设备信息");
        map.put("major","接入网");
        map.put("isCheck","是");
        map.put("businessOpenType","测试");
        list1.add(new HashMap<String,Object>(){{
            put("baseStationZ","123444");
            put("deviceTypeA","类型");
            put("deviceTypeZ","Z类型");
        }});
        list2.add(new HashMap<String,Object>(){{
            put("baseStationZ","dddaaa");
            put("deviceTypeA","A类型12332");
            put("deviceTypeZ","Z类型dadad");
        }});
        importContent1.put("normal",map);
        importContent1.put("accept",list1);
        importContent2.put("normal",map);
        importContent2.put("accept",list2);
        Map<Integer,Map<String,Object>> m = new HashMap<>();
        m.put(0,importContent1);
        m.put(1,importContent2);
        BufferedOutputStream outputStream = null;
        try {
            outputStream = new BufferedOutputStream(new FileOutputStream("H:/test/导出测试.xlsx"));
            EasyExcelUtil.complexExportMoreSheet(fileName,"导出测试",outputStream,m);
            EasyExcelUtil.complexExportMoreSheet(fileName,"导出测试",response,m);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (outputStream!=null){
                try {
                    outputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    @SneakyThrows
    @Override
    public void exportTemplate(HttpServletResponse response) {
        String templateFile="/excel/导入模板.xlsx";
        Map<Integer,String[]> fillData=new HashMap<>();
        fillData.put(5,new String[]{"江苏省","浙江省","安徽省"});
        EasyExcelUtil.exportTemplateWithFile(response,"A类销售品导入模板","导入模板",templateFile,fillData);
    }

    @Override
    public void exportSimple(HttpServletResponse response) {
        Map<Integer,String[]> map = new HashMap<>();
        map.put(0,new String[]{"123","345"});
        try {
            EasyExcelUtil.exportTemplate(response,"用户","测试",UserInfo.class,map);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
~~~

代码2：导入类-ImportServiceImpl
~~~ java
package com.zzuhkp.easyexcel.service.impl;

import com.zzuhkp.easyexcel.model.UserInfo;
import com.zzuhkp.easyexcel.service.ImportService;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.Map;

@Service
public class ImportServiceImpl implements ImportService {
    @Override
    public void importFile(List<UserInfo> list, Map<String, Object> condition) {
        list.forEach(System.out::println);
    }
}

~~~

#### 参与贡献

1.  Fork 本仓库
2.  新建 dev 分支
3.  提交代码
4.  新建 Pull Request
