package com.xiaoyu;

import com.xiaoyu.exception.InvalidParametersException;
import com.xiaoyu.utils.ExcelUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 根据匹配规则把一列生成对应的另一列,适用于有大量重复且无序的列
 */
public class ImportExcel {
    @Test
    public void testImportToBean() throws IOException, InvalidParametersException {

        Map<String, String> map = new HashMap<>();

        // 当数据转换字典用，匹配规则一定要全
        // key：规则名称 value：未解决原因
        map.put("密码管理", "配置项功能，通过参数可关闭");
        map.put("资源管理", "业务逻辑不影响");
        map.put("代码注入", "前端包含过滤器，后端参数整合校验");
        map.put("跨站脚本", "配置过滤器，防止跨站脚本攻击");
        map.put("输入验证", "数据库配置非法字符过滤");
        map.put("代码质量", "业务逻辑不影响，仅系统使用");

        // 手动将代码扫描文件规则名称列复制到exportExcelByList2.xls文件第一列中
        File file  = new File("exportExcelByList2.xls");
        // 导入文件
        List list =  ExcelUtils.importExcel(file);
        // 将不同的规则名称转换成与之对应的未解决原因
        for(int i = 0; i<list.size(); i++){
            ArrayList array = (ArrayList)list.get(i);// array代表表格第i行
            for (Map.Entry<String, String> entry : map.entrySet()) {
                if(array.get(0).equals(entry.getKey())) {// 规则名称 array.get(0)代表表格第一列
                    array.set(0, entry.getValue());// 未解决原因
                }
            }
        }
        System.out.println(list.size());

        // 将对应的未解决原因生成文件exportExcelByList1.xls，后续复制到代码扫描文件未解决原因列就好了
        File f = new File("exportExcelByList1.xls");
        OutputStream out =new FileOutputStream(f);
        HSSFWorkbook sheets = ExcelUtils.exportExcel("测试导出List", list);
        sheets.write(out);
        out.flush();
        out.close();

    }
}
