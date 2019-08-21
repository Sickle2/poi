import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * @program: poi
 * @description:
 * @author: sickle
 * @create: 2019-08-19 16:23
 **/
public class Test {

    public static void main(String args[]) {
        String sheetName = "sheetname1";
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            String filePathDaily = "D:\\Desktop\\";
            String fileName = "exceltest.xlsx";

            File filePath = new File(filePathDaily);
            if (!filePath.exists() && !filePath.isDirectory()) {
                filePath.mkdir();
            }
            //1.统计昨日交易量
            File file = new File(filePathDaily + fileName);

            List<String> title = Arrays.asList("序号", "类型", "统计", "合计：交易量", "合计：收入",
                    "2019-01-01", "2019-01-02", "2019-01-03", "2019-01-04", "2019-01-05",
                    "2019-01-01收入", "2019-01-02收入", "2019-01-03收入", "2019-01-04收入", "2019-01-05收入");
            List<String> titleStyle = Arrays.asList("int", "String", "String", "int", "double",
                    "int", "int", "int", "int", "int",
                    "double", "double", "double", "double", "double");
            List<String> dates = Arrays.asList("2019-01-01", "2019-01-02", "2019-01-03", "2019-01-04", "2019-01-05");

            Map<String, List<Object>> day2ColValueList = new LinkedHashMap<String, List<Object>>();
//            day2ColValueList.put("类型1业务1", Arrays.<Object>asList("类型1", "业务1", 500, 1000.00, 100, 50, 100, 150, 100, 200.00, 150.00, 200.00, 250.00, 200.00));
//            day2ColValueList.put("类型1业务2", Arrays.<Object>asList("类型1", "业务2", 400, 800.00, 80, 60, 80, 100, 80, 160.00, 130.00, 160.00, 190.00, 160.00));
//            day2ColValueList.put("类型2业务1", Arrays.<Object>asList("类型2", "业务1", 600, 1200.00, 120, 70, 120, 170, 120, 240.00, 170.00, 240.00, 290.00, 240.00));
//            day2ColValueList.put("类型2业务2", Arrays.<Object>asList("类型2", "业务2", 200, 400.00, 40, 140, 40, 80, 40, 80.00, 100.00, 80.00, 60.00, 80.00));

            day2ColValueList.put("类型1邀请人数", Arrays.<Object>asList("类型1", "邀请人数", 410));
            day2ColValueList.put("类型1未邀请人数", Arrays.<Object>asList("类型1", "未邀请人数", 410-410));
            day2ColValueList.put("类型1注册人数", Arrays.<Object>asList("类型1", "注册人数", 404));
            day2ColValueList.put("类型1未注册人数", Arrays.<Object>asList("类型1", "未注册人数", 410-404));
            day2ColValueList.put("类型1住宿人数", Arrays.<Object>asList("类型1", "住宿人数", 355));
            day2ColValueList.put("类型1未住宿人数", Arrays.<Object>asList("类型1", "未住宿人数", 410-355));
            day2ColValueList.put("类型1接机人数", Arrays.<Object>asList("类型1", "接机人数", 95));
            day2ColValueList.put("类型1未接机人数", Arrays.<Object>asList("类型1", "未接机人数", 410-95));
//            day2ColValueList.put("类型2业务1", Arrays.<Object>asList("总人数", "住宿人数", 600, 1200.00, 120, 70, 120, 170, 120, 240.00, 170.00, 240.00, 290.00, 240.00));
//            day2ColValueList.put("类型2业务2", Arrays.<Object>asList("总人数", "接机人数", 200, 400.00, 40, 140, 40, 80, 40, 80.00, 100.00, 80.00, 60.00, 80.00));

            MyExcleChart2.doWork(title, titleStyle, day2ColValueList, file, sheetName, wb, dates.size());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
