import com.alibaba.fastjson.JSONArray;

/**
 * @description:
 * @author: nianhen
 * @since: 2022-11-27 10:23
 **/
public class PPTGeneratorTest {
    public static void main(String[] args) {
        PPTGenerator pptGenerator = PPTGenerator.getInstance();
        // 三组
        JSONArray tableData = getTableData();
        String filePath = pptGenerator.generate(tableData, "test");
        System.out.printf("三组数据，PPT路径为: %s 。", filePath);
        // 两组
        tableData.remove(0);
        String filePath1 = pptGenerator.generate(tableData, "test1.pptx");
        System.out.printf("两组数据，PPT路径为: %s 。", filePath1);
        // 一组
        tableData.remove(0);
        String filePath2 = pptGenerator.generate(tableData, "test2.pptx");
        System.out.printf("一组数据，PPT路径为: %s 。", filePath2);

    }

    private static JSONArray getTableData() {
        return JSONArray.parseArray("[\n" +
                "  [\n" +
                "    {\n" +
                "      \"name\": \"Line 1\",\n" +
                "      \"num\": 10,\n" +
                "      \"children\": [\n" +
                "        {\n" +
                "          \"name\": \"Area 11\",\n" +
                "          \"num\": 9\n" +
                "        },\n" +
                "        {\n" +
                "          \"name\": \"Area 12\",\n" +
                "          \"num\": 1\n" +
                "        }\n" +
                "      ]\n" +
                "    },\n" +
                "    {\n" +
                "      \"name\": \"Line 2\",\n" +
                "      \"num\": 20,\n" +
                "      \"children\": [\n" +
                "        {\n" +
                "          \"name\": \"Area 21\",\n" +
                "          \"num\": 19\n" +
                "        },\n" +
                "        {\n" +
                "          \"name\": \"Area 22\",\n" +
                "          \"num\": 1\n" +
                "        }\n" +
                "      ]\n" +
                "    }\n" +
                "  ],\n" +
                "  [\n" +
                "    {\n" +
                "      \"name\": \"Line 3\",\n" +
                "      \"num\": 10,\n" +
                "      \"children\": [\n" +
                "        {\n" +
                "          \"name\": \"Area 31\",\n" +
                "          \"num\": 9\n" +
                "        },\n" +
                "        {\n" +
                "          \"name\": \"Area 32\",\n" +
                "          \"num\": 1\n" +
                "        }\n" +
                "      ]\n" +
                "    }\n" +
                "  ],\n" +
                "  [\n" +
                "    {\n" +
                "      \"name\": \"Line 4\",\n" +
                "      \"num\": 10,\n" +
                "      \"children\": [\n" +
                "        {\n" +
                "          \"name\": \"Area 41\",\n" +
                "          \"num\": 5\n" +
                "        },\n" +
                "        {\n" +
                "          \"name\": \"Area 42\",\n" +
                "          \"num\": 5\n" +
                "        }\n" +
                "      ]\n" +
                "    }\n" +
                "  ]\n" +
                "]\n" +
                "\n" +
                "\n");
    }
}
