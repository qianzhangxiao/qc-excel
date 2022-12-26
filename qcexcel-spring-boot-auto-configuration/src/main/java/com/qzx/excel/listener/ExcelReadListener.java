package com.qzx.excel.listener;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.qzx.excel.annotation.ExcelNoJudge;
import com.qzx.excel.excel.ExcelException;
import com.qzx.excel.utils.SpringBootBeanUtil;
import lombok.extern.slf4j.Slf4j;
import org.springframework.context.ApplicationContext;
import org.springframework.util.ObjectUtils;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * @description: excel导入监听类，需要为单sheet页，多sheet页会导致数据导入重复
 * @author: qc
 * @time: 2021/5/25 11:14
 */
@Slf4j
public class ExcelReadListener<T> extends AnalysisEventListener<T> {

    private Class<?> s;

    private Class<T> t;
    /**
     * 调用的service方法名
     * */
    private String methodName;

    /**
     * 反射调用方法的其他条件，使用map封装
     * */
    private Map<String,Object> map;

    /**
     *
     * @param s 被调用方法的对象（需要被spring托管）
     * @param methodName 被调用的方法名（public修饰）
     * @param map 调用方法的额外条件
     */
    public ExcelReadListener(Class<?> s, String methodName, Map<String,Object> map,Class<T> t){
        this.s=s;
        this.methodName=methodName;
        this.map=map;
        this.t = t;
    }
    /**
     * 每隔5条存储数据库，实际使用中可以3000条，然后清理list ，方便内存回收
     */
    private static final int BATCH_COUNT = 5;
    private AtomicInteger totalCount = new AtomicInteger(0);
    List<T> list = new ArrayList<>();


    /**
     * 判断模板是否符合要求
     * @param headMap
     * @param context
     */
    @Override
    public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
        // 获取数据实体的字段列表
        Field[] fields = t.getDeclaredFields();
        // 遍历字段进行判断
        for (Field field : fields) {
            // 获取当前字段上的ExcelProperty注解信息
            ExcelProperty fieldAnnotation = field.getAnnotation(ExcelProperty.class);
            // 获取标注了ExcelNoJudge的字段
            ExcelNoJudge judge = field.getAnnotation(ExcelNoJudge.class);
            // 判断当前字段上是否存在ExcelProperty、ExcelNoJudge注解
            if (fieldAnnotation != null&&judge==null) {
                // 存在ExcelProperty注解则根据注解的index索引到表头中获取对应的表头名
                String headName = headMap.get(fieldAnnotation.index());
                // 判断表头是否为空或是否和当前字段设置的表头名不相同
                if (ObjectUtils.isEmpty(headName) || !headName.equals(fieldAnnotation.value()[0])) {
                    // 如果为空或不相同，则抛出异常不再往下执行
                    throw new ExcelException("模板错误，请检查导入模板");
                }
            }
        }
    }

    /*这个每一条数据解析都会来调用*/
    @Override
    public void invoke(T t, AnalysisContext analysisContext) {
        try{
            totalCount.incrementAndGet();
            list.add(t);
            if(list.size()>=BATCH_COUNT){
                invokeMethod();
                // 存储完成清理 list
                list.clear();
            }
        }catch (Exception e){
            e.printStackTrace();
            throw new ExcelException(e.getMessage());
        }
    }

    /*所有数据解析完成了 都会来调用*/
    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        try{
            if (totalCount.intValue()==0){
                throw new ExcelException("模板不正确或者未填写信息，请确认");
            }
            invokeMethod();
            log.info("解析完所有数据,共导入{}条数据",totalCount.intValue());
        }catch (Exception e){
            e.printStackTrace();
            log.info("解析数据出错");
            throw new ExcelException(e.getMessage());
        }
    }

    public void invokeMethod() throws ClassNotFoundException, NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        //从ApplicationContext中取出已创建好的的对象
        //不可直接反射创建service对象，因为反射创建出来的对象无法实例化dao接口
        ApplicationContext applicationContext = SpringBootBeanUtil.getApplicationContext();
        //反射创建service实体对象，和实体类
        Class<?> serviceImplType = this.s;
        //反射设置方法参数。
        Method method = serviceImplType.getDeclaredMethod(this.methodName,List.class,Map.class);
        //在ApplicationContext中根据class取出已实例化的bean
        method.invoke(applicationContext.getBean(serviceImplType),this.list,this.map);
    }
}
