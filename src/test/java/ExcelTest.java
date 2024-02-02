import org.junit.jupiter.api.Test;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import java.io.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelTest {

    public static int COUNT = 1000;

    /**
     * if the template file has mergedRegions,it is really slow.
     * And will be slower and slower.
     */
    @Test
    public void slowTest() {
        List<Map<String, String>> mapList = new ArrayList<>();
        for (int i = 0; i < COUNT; i++) {
            Map<String, String> map = new HashMap<>();
            map.put("mergeA", "mergeA" + i);
            map.put("mergeB", "mergeB" + i);
            map.put("mergeC", "mergeC" + i);
            map.put("d", "d" + i);
            map.put("e", "e" + i);
            mapList.add(map);
        }

        HashMap<String, Object> map = new HashMap<>();
        map.put("rowList", mapList);
        byte[] bytes = getBytesFromTemplate(map, "ExcelTemplateSlow.xlsx");
        saveBytesToLocalFile(bytes, "ExcelExportSlow.xlsx");
    }

    /**
     * It will be normal when there is no mergedRegions.
     */
    @Test
    public void fastTest() {
        List<Map<String, String>> mapList = new ArrayList<>();
        for (int i = 0; i < COUNT; i++) {
            Map<String, String> map = new HashMap<>();
            map.put("a", "a" + i);
            map.put("b", "b" + i);
            map.put("c", "c" + i);
            map.put("d", "d" + i);
            map.put("e", "e" + i);
            mapList.add(map);
        }

        HashMap<String, Object> map = new HashMap<>();
        map.put("rowList", mapList);
        byte[] bytes = getBytesFromTemplate(map, "ExcelTemplateFast.xlsx");
        saveBytesToLocalFile(bytes, "ExcelExportFast.xlsx");
    }

    static public void saveBytesToLocalFile(byte[] bytes, String filename) {
        FileOutputStream fos = null;
        try {
            String filePath = "ExcelTemp";
            File path = new File(filePath);
            if (!path.exists() && !path.isDirectory()) {
                path.mkdir();
            }
            LocalDateTime now = LocalDateTime.now();
            DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss-SSS");
            filename = now.format(dtf) + filename;

            fos = new FileOutputStream(filePath + "/" + filename);
            fos.write(bytes);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (fos != null) {
                try {
                    fos.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    static public byte[] getBytesFromTemplate(Map<String, Object> map, String templateFileName) {
        try {
            InputStream is = ExcelTest.class.getClassLoader().getResourceAsStream(templateFileName);
            ByteArrayOutputStream os = new ByteArrayOutputStream();
            Context context = new Context(map);
            JxlsHelper.getInstance().processTemplate(is, os, context);
            if (is != null) {
                is.close();
            }
            return os.toByteArray();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return "error".getBytes();
    }

}
