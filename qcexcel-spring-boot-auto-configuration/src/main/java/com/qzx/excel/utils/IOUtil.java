package com.qzx.excel.utils;

import org.springframework.core.io.ClassPathResource;

import java.io.*;

/**
 * @description: 文件操作工具类
 * @author: qc
 * @time: 2020/8/25 19:55
 */
public class IOUtil {
    private static InputStream is;
    private static BufferedInputStream bis;

    private static OutputStream os;
    private static BufferedOutputStream bos;
    /**
     *
     *
     * @description: 创建文件目录
      * @param path:文件路径
     * @return: 是否创建成功
     * @author: qc
     * @time: 2020/8/25 19:55
     */
    public static boolean createDirectory(String path){

        File file=getFile(path);
        if(!file.exists()){
            file.mkdir();
            return true;
        }
        return false;
    }

    public static File getFile(String path){
        File file=new File(path);
        return file;
    }

    /**
     * 获取文件输入流
     * @param path；文件路径
     * @param fileName：文件名称
     * @return
     */
    public static InputStream getInputStream(String path,String fileName){
        return getInputStreamFromClassPath(getPath(path)+fileName);
    }

    /**
     * 获取文件输入流
     * @param file：同时包括文件路径和文件名
     * @return
     */
    public static InputStream getInputStream(String file){
        try {
            is= new FileInputStream(getFile(file));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return is;
    }
    /**
     * 获取文件输出流，
     * @param path：文件路径
     * @param fileName：文件名
     * @return
     */
    public static OutputStream getOutputStream(String path,String fileName){
        return getOutputStream(getPath(path)+fileName);
    }

    /**
     * 获取文件输出流，同时包括文件路径和文件名
     * @param file
     * @return
     */
    public static OutputStream getOutputStream(String file){
        try {
            os= new FileOutputStream(getFile(file));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return os;
    }
    public static BufferedInputStream getBufferInput(InputStream inputStream){
        bis=new BufferedInputStream(inputStream);
        return bis;
    }
    public static BufferedInputStream getBufferInput(String path,String fileName){
        return getBufferInput(getPath(path)+fileName);
    }
    public static BufferedInputStream getBufferInput(String file){
        bis=new BufferedInputStream(getInputStream(file));
        return bis;
    }
    public static BufferedInputStream getBufferInputFromClassPath(String path,String fileName){
        return getBufferInputFromClassPath(getPath(path)+fileName);
    }


    public static BufferedInputStream getBufferInputFromClassPath(String file){
        bis=new BufferedInputStream(getInputStreamFromClassPath(file));
        return bis;
    }
    public static BufferedOutputStream getBufferOutput(String path,String fileName){
        bos=new BufferedOutputStream(getOutputStream(path,fileName));
        return bos;
    }

    /**
     * 获取classpath下的文件，需给出路径和文件名
     * @param path:/public、public、/public/、public/
     * @param fileName:0a2ebee9-e84f-4900-8bb2-2406be9c65a6.jpg
     * @return
     */
     public static InputStream getInputStreamFromClassPath(String path,String fileName){
         return getInputStreamFromClassPath(getPath(path)+fileName);
     }

    /**
     * 从resource获取文件，需指定文件路径及名称，例如：/public/0a2ebee9-e84f-4900-8bb2-2406be9c65a6.jpg
     * @param file
     * @return
     */
    public static InputStream getInputStreamFromClassPath(String file){
        ClassPathResource classPathResource = new ClassPathResource(file);
        InputStream inputStream =null;
        try {
            inputStream= classPathResource.getInputStream();

        } catch (IOException e) {
            e.printStackTrace();
        }
        return inputStream;
    }
    private static String getPath(String path){
        if(path.endsWith("/")||path.endsWith("\\")){
            return path;
        }
        return path+"/";
    }

    public static void closeIo(){
        try {
            if(is!=null){
                is.close();
            }
            if(bis!=null){
                bis.close();
            }
            if(os!=null){
                os.close();
            }
            if(bos!=null){
                bos.close();
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
