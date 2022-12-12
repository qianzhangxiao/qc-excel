package com.qzx.excel.utils;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.enums.WriteDirectionEnum;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.alibaba.excel.write.metadata.fill.FillWrapper;
import com.qzx.excel.listener.ExcelReadListener;
import lombok.extern.slf4j.Slf4j;
import org.springframework.util.ObjectUtils;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * @description: easyExcel工具类
 * @author: qc
 * @time: 2021/5/24 14:01
 */
@Slf4j
public class EasyExcelUtil {

    /*私有构造器*/
   private EasyExcelUtil(){}

    /**
     *
     *
     * @description: 简单导出excel文件
      * @param sheetMap 封装的excel数据，Map的String为sheet页名称，List<T>为具体的数据
     * @param fileName 导出的文件名
     * @param excludeMap:需排除的表头，key和sheetMap的key一致
     * @return:
     * @author: qc
     * @time: 2021/5/24 14:24
     */
   /*导出Excel文件*/
    public static <T> boolean writeExcel(HttpServletResponse response,
                                         Map<String,List<T>> sheetMap,
                                         String fileName,
                                         Class<T> tClass,
                                         Map<String,List<String>> excludeMap){
        boolean b=false;
        log.info("导出Excel工具类");
        ExcelWriter excelWriter=null;
        try{
            /*设置response格式*/
            response.setContentType("application/vnd.ms-excel");
            response.setCharacterEncoding("utf-8");
            fileName = URLEncoder.encode(fileName, "UTF-8");
            /*设置导出excel格式为xlsx*/
            response.setHeader("Content-disposition", "attachment;filename="+fileName+".xlsx");
            /*获取EasyExcel对象*/
            excelWriter = EasyExcel.write(response.getOutputStream()).build();
            /*获取sheet页页码*/
            AtomicInteger i= new AtomicInteger();
            ExcelWriter finalExcelWriter = excelWriter;
            /*遍历sheet页*/
            sheetMap.forEach((k,v)->{
                /*获取排除的表头*/
                /*获取排除的表头*/
                List<String> excludeList=null;
                if(!ObjectUtils.isEmpty(excludeMap)){
                    excludeList=excludeMap.get(k);
                }
                /*按sheet页写入数据*/
                WriteSheet writeSheet=EasyExcel.writerSheet(i.getAndIncrement(), k).head(tClass).excludeColumnFieldNames(excludeList).build();
                finalExcelWriter.write(v,writeSheet);
            });
            b=true;
            log.info("导出Excel工具类-成功");
        }catch (Exception e){
            e.printStackTrace();
            log.info("导出Excel工具类-异常");
            b=false;
        }finally {
            if(excelWriter!=null){
                excelWriter.finish();
            }
        }
        return b;
    }

    /**
     *
     *
     * @description: 对导入的excel文件进行操作
      * @param file：读取的excel文件
     * @param <T> 实体类
     * @param s service类
     * @param methodName:service对应的方法
     * @param condition 调用监听器额外的条件
     * @return: 是否成功
     * @author: qc
     * @time: 2021/5/25 14:01
     */
    public static <T> boolean importExcel(InputStream file,Class<?> s,String methodName,Class<T> tClass,Map<String,Object> condition){
        try{
            log.info("开始解析excel文件");
            // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
            // EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet().doRead();
            /*doReadAll()读取所有sheet页*/
            EasyExcel.read(file,tClass,new ExcelReadListener<T>(s,methodName,condition)).doReadAll();
            log.info("解析excel文件成功");
            return true;
        }catch (Exception e){
            e.printStackTrace();
            log.error("解析excel文件异常");
            return false;
        }finally {
            if(file!=null){
                try {
                    file.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
    public static <T> boolean importExcel(String fileName,Class<?> s,String methodName,Class<T> tClass,Map<String,Object> condition){
        try{
            log.info("开始解析excel文件");
            // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
            // EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet().doRead();
            /*doReadAll()读取所有sheet页*/
            EasyExcel.read(fileName,tClass,new ExcelReadListener<T>(s,methodName,condition)).doReadAll();
            log.info("解析excel文件成功");
            return false;
        }catch (Exception e){
            e.printStackTrace();
            log.error("解析excel文件异常");
            return true;
        }
    }

    /*浏览器下载文件*/
    public static OutputStream getOutPutStream(HttpServletResponse response,String fileName) throws IOException {
        String exportFileName=URLEncoder.encode(fileName+ ExcelTypeEnum.XLS.getValue(), StandardCharsets.UTF_8.toString());
        response.setContentType("application/force-download");
        response.setHeader("Content-Disposition","attachment;fileName="+exportFileName);
        return response.getOutputStream();
    }

    /**
     *  @description:按模板导出
     * @param t 导出模板的实体类
     * @param map 需要添加下拉菜单的列，列从0开始,value-->String[]为具体下拉菜单的值
     */
    public static<T> void exportTemplate(HttpServletResponse response,String templateName,String sheetName,T t,Map<Integer,String[]> map) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        String fileName = URLEncoder.encode(templateName, "UTF-8");
        response.setHeader("content-disposition", "attachment;filename=" + fileName + ".xlsx");
        EasyExcel.write(response.getOutputStream(), t.getClass())
                .registerWriteHandler(new SelfWriteHandle(map))
                .sheet(sheetName)
                .doWrite(new ArrayList<>());
    }

    /**
     * 复杂模板导出
     * @param importFile 导入模板，包含路径及名称，如/static/test.xlsx
     * @param exportFileName 导出的文件名称
     * @param sheetName sheet名称
     * @param response 返回响应
     * @param importContent 模板所需数据
     */
    public static boolean complexExport(String importFile,String exportFileName,String sheetName,HttpServletResponse response,Map<String,Object> importContent){
        InputStream file=IOUtil.getInputStreamFromClassPath(importFile);
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        ExcelWriter writer= null;
        try {
            writer = EasyExcel.write(EasyExcelUtil.getOutPutStream(response,exportFileName)).withTemplate(file).build();
            WriteSheet writeSheet=EasyExcel.writerSheet(0,sheetName).registerWriteHandler(new MyHeader()).build();
            FillConfig fillConfig=FillConfig.builder().direction(WriteDirectionEnum.VERTICAL).forceNewRow(Boolean.TRUE).build();
            ExcelWriter finalWriter = writer;
            importContent.forEach((k, v)->{
                if (v instanceof Map){
                    finalWriter.fill(v,fillConfig,writeSheet);
                }
                if (v instanceof Collection<?>){
                    finalWriter.fill(new FillWrapper(k, (Collection<?>)v ),fillConfig,writeSheet);
                }
            });
            writer.finish();
            return true;
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }

    }
}
