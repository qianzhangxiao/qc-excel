package com.qzx.excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.qzx.excel.utils.SpringBootBeanUtil;
import lombok.extern.slf4j.Slf4j;
import org.springframework.context.ApplicationContext;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @description:
 * @author: qc
 * @time: 2021/5/25 11:14
 */
@Slf4j
public class ExcelReadListener<T> extends AnalysisEventListener<T> {

    private Class<?> s;
    /*调用的service方法名*/
    private String methodName;

    /*反射调用方法的其他条件，使用map封装*/
    private Map<String,Object> map;

    public ExcelReadListener(Class<?> s, String methodName, Map<String,Object> map){
        this.s=s;
        this.methodName=methodName;
        this.map=map;
    }
    /**
     * 每隔5条存储数据库，实际使用中可以3000条，然后清理list ，方便内存回收
     */
    private static final int BATCH_COUNT = 5;
    private volatile int totalCount = 0;
    List<T> list = new ArrayList<>();


    /*这个每一条数据解析都会来调用*/
    @Override
    public void invoke(T t, AnalysisContext analysisContext) {
        try{
            list.add(t);
            if(list.size()>=BATCH_COUNT){
                invokeMethod();
                // 存储完成清理 list
                list.clear();
            }
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    /*所有数据解析完成了 都会来调用*/
    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        try{
            invokeMethod();
            log.info("解析完所有数据");
        }catch (Exception e){
            e.printStackTrace();
            log.info("解析数据出错");
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
