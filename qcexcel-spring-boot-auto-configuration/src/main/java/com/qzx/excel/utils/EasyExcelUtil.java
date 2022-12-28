package com.qzx.excel.utils;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.enums.WriteDirectionEnum;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.alibaba.excel.write.metadata.fill.FillWrapper;
import com.qzx.excel.annotation.ExcelNoJudge;
import com.qzx.excel.excel.ExcelException;
import com.qzx.excel.listener.ExcelReadListener;
import lombok.extern.slf4j.Slf4j;
import org.springframework.util.ObjectUtils;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;
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
     * @param tClass 模板类
     * @param condition 调用监听器额外的条件
     * @param templateHeaderMap 模板表头（从0开始，用来校验上传文件与所需模板是否匹配）
     * @return: 是否成功
     * @author: qc
     * @time: 2021/5/25 14:01
     */
    public static <T> boolean importExcel(InputStream file,Class<?> s,String methodName,Class<T> tClass,Map<String,Object> condition,Map<Integer,String > templateHeaderMap) throws Exception {
        try{
            log.info("开始解析excel文件");
            // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
            // EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet().doRead();
            /*doReadAll()读取所有sheet页*/
            EasyExcel.read(file,tClass,new ExcelReadListener<T>(s,methodName,condition,tClass,templateHeaderMap)).autoCloseStream(true).doReadAll();
            log.info("解析excel文件成功");
            return true;
        }catch (ExcelException e){
          log.error("导入模板有误");
          throw new ExcelException(e.getMessage());
        } catch (Exception e){
            e.printStackTrace();
            log.error("解析excel文件异常");
            throw new Exception("文件解析异常，只支持xlsx/xsl格式");
        }
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
    public static <T> boolean importExcel(InputStream file,Class<?> s,String methodName,Class<T> tClass,Map<String,Object> condition) throws Exception {
        try{
            log.info("开始解析excel文件");
            // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
            // EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet().doRead();
            /*doReadAll()读取所有sheet页*/
            EasyExcel.read(file,tClass,new ExcelReadListener<T>(s,methodName,condition,tClass)).autoCloseStream(true).doReadAll();
            log.info("解析excel文件成功");
            return true;
        }catch (ExcelException e){
            log.error("导入模板有误");
            throw new ExcelException(e.getMessage());
        } catch (Exception e){
            e.printStackTrace();
            log.error("解析excel文件异常");
            throw new Exception("文件解析异常，只支持xlsx/xsl格式");
        }
    }
    /**
     * 通过文件名导入excel
     * @param fileName 文件名
     * @param s 被调用处理方法的类对象（需要被spring托管）
     * @param methodName 被调用处理方法
     * @param tClass 模板类
     * @param condition 处理条件
     * @throws Exception
     */
    public static <T> boolean importExcel(String fileName,Class<?> s,String methodName,Class<T> tClass,Map<String,Object> condition) throws Exception {
        try{
            log.info("开始解析excel文件");
            // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
            // EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet().doRead();
            /*doReadAll()读取所有sheet页*/
            EasyExcel.read(fileName,tClass,new ExcelReadListener<T>(s,methodName,condition,tClass)).autoCloseStream(true).doReadAll();
            log.info("解析excel文件成功");
            return true;
        }catch (ExcelException e){
            log.error("导入模板有误");
            throw new ExcelException(e.getMessage());
        } catch (Exception e){
            e.printStackTrace();
            log.error("解析excel文件异常");
            throw new Exception("文件解析异常，只支持xlsx/xsl格式");
        }
    }

    /**
     * 通过文件名导入excel
     * @param fileName 文件名
     * @param s 被调用处理方法的类对象（需要被spring托管）
     * @param methodName 被调用处理方法
     * @param tClass 模板类
     * @param condition 处理条件
     * @param templateHeaderMap 模板头（从0开始，用来校验上传文件与所需模板是否匹配）
     * @throws Exception
     */
    public static <T> boolean importExcel(String fileName,Class<?> s,String methodName,Class<T> tClass,Map<String,Object> condition,Map<Integer,String> templateHeaderMap) throws Exception {
        try{
            log.info("开始解析excel文件");
            // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
            // EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet().doRead();
            /*doReadAll()读取所有sheet页*/
            EasyExcel.read(fileName,tClass,new ExcelReadListener<T>(s,methodName,condition,tClass,templateHeaderMap)).autoCloseStream(true).doReadAll();
            log.info("解析excel文件成功");
            return true;
        }catch (ExcelException e){
            log.error("导入模板有误");
            throw new ExcelException(e.getMessage());
        } catch (Exception e){
            e.printStackTrace();
            log.error("解析excel文件异常");
            throw new Exception("文件解析异常，只支持xlsx/xsl格式");
        }
    }

    /*浏览器下载文件*/
    public static OutputStream getOutPutStream(HttpServletResponse response,String fileName) throws IOException {
        String exportFileName=URLEncoder.encode(fileName+ ExcelTypeEnum.XLSX.getValue(), StandardCharsets.UTF_8.toString());
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        response.setHeader("Content-Disposition","attachment;fileName="+exportFileName);
        return response.getOutputStream();
    }

    /**
     *  @description:按模板导出
     * @param t 导出模板的实体类
     * @param map 需要添加下拉菜单的列，列从0开始,value-->String[]为具体下拉菜单的值
     */
    public static<T> void exportTemplate(HttpServletResponse response,String templateName,String sheetName,Class<T> tClass,Map<Integer,String[]> map) throws IOException {
        EasyExcel.write(getOutPutStream(response,templateName), tClass)
                .registerWriteHandler(new SelfWriteHandle(map))
                .sheet(sheetName)
                .doWrite(new ArrayList<>());
    }

    /**
     * 按文件模板导出
     * @param response
     * @param templateName 导出模板名称
     * @param sheetName sheet名称
     * @param templateStream 模板文件数据流
     * @param fillData 填充数据
     */
    public static void exportTemplateWithFile(HttpServletResponse response,String templateName,String sheetName,
                                              InputStream templateStream,Map<Integer,String[]> fillData) throws IOException {
        EasyExcel.write(getOutPutStream(response,templateName))
                .withTemplate(templateStream)
                .autoCloseStream(true) //自动关闭流
                .registerWriteHandler(new SelfWriteHandle(fillData))
                .sheet(sheetName)
                .doWrite(new ArrayList<>());
    }
    /**
     * 按文件模板导出
     * @param response
     * @param templateName 导出模板名称
     * @param sheetName sheet名称
     * @param templateFile 模板文件
     * @param fillData 填充数据
     */
    public static void exportTemplateWithFile(HttpServletResponse response,String templateName,String sheetName,
                                              String templateFile,Map<Integer,String[]> fillData) throws IOException {
        InputStream templateStream=IOUtil.getInputStreamFromClassPath(templateFile);
        exportTemplateWithFile(response,templateName,sheetName,templateStream,fillData);
    }
    /**
     * 单sheet页复杂模板导出
     * @param importFile 导入模板，包含路径及名称，如/static/test.xlsx
     * @param exportFileName 导出的文件名称
     * @param sheetName sheet名称
     * @param response 返回响应
     * @param importContent 模板所需数据
     */
    public static void complexExport(InputStream importFile,String exportFileName,String sheetName,HttpServletResponse response,Map<String,Object> importContent) throws IOException {
        complexExport(importFile,exportFileName,sheetName,getOutPutStream(response,exportFileName),importContent);
    }
    /**
     * 单sheet页复杂模板导出
     * @param importFile 导入模板，包含路径及名称，如/static/test.xlsx
     * @param exportFileName 导出的文件名称
     * @param sheetName sheet名称
     * @param outputStream 输出流
     * @param importContent 模板所需数据
     */
    public static void complexExport(InputStream importFile,String exportFileName,String sheetName,OutputStream outputStream,Map<String,Object> importContent){
        ExcelWriter writer= null;
        writer = EasyExcel.write(outputStream).withTemplate(importFile).autoCloseStream(true).build();
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
    }

    /**
     * 多sheet页复杂模板导出
     * @param importFile 导入模板，包含路径及名称，如/static/test.xlsx
     * @param exportFileName 导出的文件名称
     * @param outputStream 输出流
     * @param importContent key为sheet页码，从0开始，value为sheet页填充所需数据
     */
    public static void complexExportMoreSheet(InputStream importFile,String exportFileName,OutputStream outputStream,Map<Integer,Map<String,Object>> importContent){
        ExcelWriter writer= null;
        writer = EasyExcel.write(outputStream).withTemplate(importFile).autoCloseStream(true).build();
        ExcelWriter finalWriter = writer;
        importContent.forEach((sheetNo, sheetContent)->{
            WriteSheet writeSheet=EasyExcel.writerSheet(sheetNo).registerWriteHandler(new MyHeader()).build();
            FillConfig fillConfig=FillConfig.builder().direction(WriteDirectionEnum.VERTICAL).forceNewRow(Boolean.TRUE).build();
            sheetContent.forEach((k, v)->{
                if (v instanceof Map){
                    finalWriter.fill(v,fillConfig,writeSheet);
                }
                if (v instanceof Collection<?>){
                    finalWriter.fill(new FillWrapper(k, (Collection<?>)v ),fillConfig,writeSheet);
                }
            });
        });
        writer.finish();
    }

    /**
     * 多sheet页复杂模板导出
     * @param importFile 导入模板，包含路径及名称，如/static/test.xlsx
     * @param exportFileName 导出的文件名称
     * @param response 输出流
     * @param importContent key为sheet页码，从0开始，value为sheet页填充所需数据
     */
    public static void complexExportMoreSheet(InputStream importFile,String exportFileName,HttpServletResponse response,Map<Integer,Map<String,Object>> importContent) throws IOException {
        complexExportMoreSheet(importFile,exportFileName,getOutPutStream(response,exportFileName),importContent);
    }

    /**
     * 单sheet页复杂模板导出，模板位于项目classpath环境下
     * @param importFile 导入模板，包含路径及名称，如/static/test.xlsx
     * @param exportFileName 导出的文件名称
     * @param sheetName sheet名称
     * @param response 返回响应
     * @param importContent 模板所需数据
     */
    public static void complexExport(String importFile,String exportFileName,String sheetName,HttpServletResponse response,Map<String,Object> importContent) throws IOException {
        InputStream file=IOUtil.getInputStreamFromClassPath(importFile);
        complexExport(file,exportFileName,sheetName,response,importContent);
    }
    /**
     * 单sheet页复杂模板导出，模板位于项目classpath环境下
     * @param importFile 导入模板，包含路径及名称，如/static/test.xlsx
     * @param exportFileName 导出的文件名称
     * @param sheetName sheet名称
     * @param outputStream 输出流
     * @param importContent 模板所需数据
     */
    public static void complexExport(String importFile,String exportFileName,String sheetName,OutputStream outputStream,Map<String,Object> importContent){
        InputStream file=IOUtil.getInputStreamFromClassPath(importFile);
        complexExport(file,exportFileName,sheetName,outputStream,importContent);
    }
    /**
     * 多sheet页复杂模板导出，模板位于项目classpath环境下
     * @param importFile 导入模板，包含路径及名称，如/static/test.xlsx
     * @param exportFileName 导出的文件名称
     * @param response 返回响应
     * @param importContent key为sheet页码，从0开始，value为sheet页填充所需数据
     */
    public static void complexExportMoreSheet(String importFile,String exportFileName,HttpServletResponse response,Map<Integer,Map<String,Object>> importContent) throws IOException {
        InputStream file=IOUtil.getInputStreamFromClassPath(importFile);
        complexExportMoreSheet(file,exportFileName,response,importContent);
    }
    /**
     * 多sheet页复杂模板导出，模板位于项目classpath环境下
     * @param importFile 导入模板，包含路径及名称，如/static/test.xlsx
     * @param exportFileName 导出的文件名称
     * @param outputStream 输出流
     * @param importContent key为sheet页码，从0开始，value为sheet页填充所需数据
     */
    public static void complexExportMoreSheet(String importFile,String exportFileName,OutputStream outputStream,Map<Integer,Map<String,Object>> importContent){
        InputStream file=IOUtil.getInputStreamFromClassPath(importFile);
        complexExportMoreSheet(file,exportFileName,outputStream,importContent);
    }

    /**
     * 获取指定模板的含有ExcelProperty的属性并封装成map
     * @param t
     * @return
     */
    public static Map<Integer,String> headMap(Class<?> t){
        Map<Integer,String > headMap = new ConcurrentHashMap<>();
        final Field[] fields = t.getDeclaredFields();
        for (Field field : fields) {
            final ExcelProperty property = field.getAnnotation(ExcelProperty.class);
            if (property!=null){
                headMap.put(property.index(),property.value()[0]);
            }
        }
        return headMap;
    }
    /**
     * 获取指定模板的含有ExcelProperty且不含有ExcelNoJudge注解的属性并封装成map
     * @param t
     * @return
     */
    public static Map<Integer,String> headMapWithoutNoJudge(Class<?> t){
        Map<Integer,String > headMap = new ConcurrentHashMap<>();
        final Field[] fields = t.getDeclaredFields();
        for (Field field : fields) {
            final ExcelProperty property = field.getAnnotation(ExcelProperty.class);
            final ExcelNoJudge judge = field.getAnnotation(ExcelNoJudge.class);
            if (property!=null&&judge==null){
                headMap.put(property.index(),property.value()[0]);
            }
        }
        return headMap;
    }

}
