package com.hkc.util;

import cn.hutool.captcha.CaptchaUtil;
import cn.hutool.captcha.LineCaptcha;
import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.date.DateUtil;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.lang.Console;
import cn.hutool.poi.excel.BigExcelWriter;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import com.hkc.annotation.ExcelCol;
import com.hkc.core.ExportColVo;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.Validate;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;
/**
 * @author hukunchi
 * @Title: ${file_name}
 * @Package ${package_name}
 * @Description: ${todo}
 * @date 2018/11/815:23
 */
public class XlsExportUtil {

    /**
     * 导出
     *
     * @param dirPath
     *            目录
     * @param filePath
     *            文件名
     * @param list
     *            导出的数据
     * @param sheetName
     *            excel工作本的名称
     * @param className
     *            类名
     * @throws InstantiationException
     */
    public static <E> void export(String dirPath, String filePath, List<E> list, String sheetName, Class<?> className)
            throws InstantiationException {
        File fileNew = new File(dirPath);
        if (!fileNew.isFile()) {
            fileNew.mkdirs();
        }
        File file = new File(filePath);
        OutputStream os = null;
        WritableWorkbook wwb = null;
        try {
            file.createNewFile();

            os = new FileOutputStream(file.getPath());
            wwb = Workbook.createWorkbook(os);

            WritableSheet sheet = wwb.createSheet(sheetName, 0);

            // 获取excle的所有列明、序号、对应的字段名称
            List<ExportColVo> exportColVoList = getClassAllCols(className);
            // 字段排序
            sortExcelCols(exportColVoList);

            int currRowNum = 1;
            for (int rowNum = 0; rowNum < list.size(); rowNum++) {
                E e = list.get(rowNum);

                for (int colNum = 0; colNum < exportColVoList.size(); colNum++) {

                    String cellContent = formatCellContent(e, exportColVoList.get(colNum));

                    // 标题行
                    if (rowNum == 0) {

                        Label titleLabel = new Label(colNum, rowNum, exportColVoList.get(colNum).getColName());
                        sheet.addCell(titleLabel);
                    }
                    Label dataLabel = new Label(colNum, rowNum + 1, cellContent);
                    sheet.addCell(dataLabel);
                }

                // 一万条写一次
                if (currRowNum % 10000 == 0) {
                    wwb.write();
                    os.flush();
                }
                currRowNum++;
            }

            wwb.write();
            os.flush();
        } catch (IOException | WriteException e) {
            System.out.println(e.getMessage());
        } finally {
            try {
                if (wwb != null) {
                    wwb.close();
                }
                if (os != null) {
                    os.close();
                }
            } catch (WriteException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

    /**
     * 设置可以访问
     *
     * @param method
     */
    private static void makeAccessible(Method method) {
        if ((!Modifier.isPublic(method.getModifiers()) || !Modifier.isPublic(method.getDeclaringClass().getModifiers()))
                && !method.isAccessible()) {
            method.setAccessible(true);
        }
    }

    /**
     * 获取method对象
     *
     * @param obj
     * @param methodName
     * @param parameterTypes
     * @return
     */
    private static Method getAccessibleMethod(final Object obj, final String methodName, final Class<?>... parameterTypes) {
        Validate.notNull(obj, "object can't be null");
        Validate.notBlank(methodName, "methodName can't be blank");

        for (Class<?> searchType = obj.getClass(); searchType != Object.class; searchType = searchType
                .getSuperclass()) {
            try {
                Method method = searchType.getDeclaredMethod(methodName, parameterTypes);
                makeAccessible(method);
                return method;
            } catch (NoSuchMethodException e) {
                System.out.println(e.getMessage());
                // Method不在当前类定义,继续向上转型
                continue;// new add
            }
        }
        return null;
    }

    /**
     * 反射调用对象方法
     *
     * @param obj
     * @param methodName
     * @param parameterTypes
     * @param args
     * @return
     */
    private static Object invokeMethod(final Object obj, final String methodName, final Class<?>[] parameterTypes,
                                       final Object[] args) {
        Method method = getAccessibleMethod(obj, methodName, parameterTypes);
        if (method == null) {
            throw new IllegalArgumentException("Could not find method [" + methodName + "] on target [" + obj + "]");
        }

        try {
            return method.invoke(obj, args);
        } catch (IllegalAccessException | IllegalArgumentException | InvocationTargetException e) {
            System.out.println(e.getMessage());
        }
        return null;
    }

    // 线程不安全
    private static SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");

    /**
     * 格式化单元格，支持数字、日期类型
     *
     * @param rowClass
     * @param exportColVo
     * @return
     * @throws InstantiationException
     */
    private static String formatCellContent(Object rowClass, ExportColVo exportColVo) throws InstantiationException {
        String cellContent = "";

        try {
            Object fieldValue = invokeMethod(rowClass, parGetName(exportColVo.getFieldName()), new Class[] {},
                    new Object[] {});
            String format = exportColVo.getFormat();
            if (fieldValue == null) {
                return "";
            }
            // 字符串类型直接返回
            if (fieldValue instanceof String) {
                return (String) fieldValue;
                // 数字类型，判断是否需要格式化
            } else if (fieldValue.getClass().getSuperclass().getName().equals(Number.class.getName())) {
                cellContent = fieldValue.toString();
                if (StringUtils.isNotEmpty(format)) {
                    String[] formatArr = format.trim().split(":");
                    String formatType = formatArr[0];
                    Integer decimalNum = Integer.valueOf(formatArr[1]);
                    if ("accuracy".equals(formatType)) {
                        cellContent = formatNum((Number) fieldValue, decimalNum);
                    } else if ("percent".equals(formatType)) {
                        if (decimalNum == -1) {
                            cellContent = fieldValue + "%";
                        } else {
                            cellContent = formatNum((Number) fieldValue, decimalNum) + "%";
                        }
                    }
                }
                return cellContent;
                // 判断是否是数字类型
            } else if (fieldValue instanceof Date) {
                simpleDateFormat.applyPattern(format);
                cellContent = simpleDateFormat.format((Date) fieldValue);
                return cellContent;
            }

        } catch (IllegalArgumentException e) {
            System.out.println(e.getMessage());
        }

        return cellContent;
    }

    /**
     * 获取列的排序、列名、对应的字段名称
     *
     * @param className
     * @return
     */
    private static List<ExportColVo> getClassAllCols(Class<?> className) {
        Field[] fields = className.getDeclaredFields();

        List<ExportColVo> exportColVoList = new ArrayList<ExportColVo>();

        if (fields != null && fields.length > 0) {
            for (Field field : fields) {
                ExcelCol excelCol = field.getAnnotation(ExcelCol.class);
                if (excelCol != null) {
                    Integer sort = Integer.valueOf(excelCol.sort());
                    String colName = excelCol.name();
                    String fieldName = field.getName();
                    String format = excelCol.format();
                    ExportColVo exportColVo = new ExportColVo(sort, colName, fieldName, format);
                    exportColVoList.add(exportColVo);
                }
            }
        }

        return exportColVoList;
    }

    /**
     * excel列排序
     *
     * @param exportColVoList
     */
    private static void sortExcelCols(List<ExportColVo> exportColVoList) {
        Collections.sort(exportColVoList, new Comparator<ExportColVo>() {

            @Override
            public int compare(ExportColVo o1, ExportColVo o2) {

                return o1.getSort() - o2.getSort();
            }
        });
    }

    /**
     * 格式化数字
     *
     * @param number
     * @param decimalNum
     * @return
     */
    private static String formatNum(Number number, Integer decimalNum) {
        StringBuilder formatStr = new StringBuilder("##0.");
        for (int i = 0; i < decimalNum; i++) {
            formatStr.append("0");
        }
        DecimalFormat decimalFormat = new DecimalFormat(formatStr.toString());
        return decimalFormat.format(number);
    }

    /**
     * 拼接某属性的 get方法
     *
     * @param fieldName
     * @return String
     */
    private static String parGetName(String fieldName) {
        if (null == fieldName || "".equals(fieldName)) {
            return null;
        }
        int startIndex = 0;
        if (fieldName.charAt(0) == '_')
            startIndex = 1;
        return "get" + fieldName.substring(startIndex, startIndex + 1).toUpperCase()
                + fieldName.substring(startIndex + 1);
    }

    public static int region(int min, int max) {
        Random rand = new Random();
        return rand.nextInt(max - min) + min;
    }

    public static void testXlsExport(){
        List<DemoVo> list=new ArrayList<>();
        for (int i=0;i<30;i++){
            DemoVo demoVo=new DemoVo();
            demoVo.setName("tomcat".concat(String.valueOf(i)));
            demoVo.setAge("22".concat(String.valueOf(i)));
            demoVo.setAddr("深圳".concat(String.valueOf(i)));
            demoVo.setAmout(BigDecimal.valueOf(30));
            list.add(demoVo);
        }
        // 获取临时文件存放路径
        String dir = "d://";
        // 命名规则 时间 + 两位随机数+“pawn”
        String name = DateUtil.format(new Date(), "yyyyMMddHHmmss_" + XlsExportUtil.region(10, 99));
        String fileName = name + ".xls";
        String path = dir + File.separatorChar + fileName;
        //导出EXCEL
        try {
            //写入到磁盘中
            XlsExportUtil.export(dir,path,list,fileName,DemoVo.class);
        }catch (Exception e){
            System.out.println("导出出错了");
        }
    }
    public static  void testLineCaptcha(){

        //定义图形验证码的长和宽
        LineCaptcha lineCaptcha = CaptchaUtil.createLineCaptcha(200, 100);

        //图形验证码写出，可以写出到文件，也可以写出到流
        lineCaptcha.write("d:/line2.png");
        //输出code
        System.out.println("本次生成的的验证码:"+lineCaptcha.getCode());
        //验证图形验证码的有效性，返回boolean值
        boolean result= lineCaptcha.verify("1234");
        System.out.println("result:"+result);

        //重新生成验证码
        lineCaptcha.createCode();
        lineCaptcha.write("d:/line.png");
        //新的验证码
        Console.log(lineCaptcha.getCode());
        //验证图形验证码的有效性，返回boolean值
        boolean result2=lineCaptcha.verify("1234");
        System.out.println("result2:"+result2);
    }
    public static void testBigDataXlsExport(){
        System.out.println("startTime:"+System.currentTimeMillis());
        List<DemoVo> list=new ArrayList<>();
        for (int i=0;i<300000;i++){
            DemoVo demoVo=new DemoVo();
            demoVo.setName("tomcat".concat(String.valueOf(i)));
            demoVo.setAge("22".concat(String.valueOf(i)));
            demoVo.setAddr("深圳".concat(String.valueOf(i)));
            demoVo.setAmout(BigDecimal.valueOf(30));
            list.add(demoVo);
        }
        List<DemoVo> rows = CollUtil.newArrayList(list);
        //ExcelWriter writer = ExcelUtil.getWriter("d:/writeBeanTest.xlsx");
        BigExcelWriter writer = ExcelUtil.getBigWriter("d:/sssss.xlsx");

        //自定义标题别名
        writer.addHeaderAlias("name", "姓名");
        writer.addHeaderAlias("age", "年龄");
        writer.addHeaderAlias("score", "分数");
        writer.addHeaderAlias("isPass", "是否通过");
        writer.addHeaderAlias("examDate", "考试时间");

        // 合并单元格后的标题行，使用默认标题样式
        writer.merge(4, "一班成绩单");
        // 一次性写出内容，使用默认样式
        writer.write(rows);
        // 关闭writer，释放内存
        writer.close();
        System.out.println("endTime"+System.currentTimeMillis());
    }
    public static void main(String[] args) {
        testLineCaptcha();
    }


}
