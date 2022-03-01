import cn.hutool.core.io.FileUtil;
import cn.hutool.json.JSONUtil;

import java.util.List;

public class Test {
    public static void main(String[] args) {
        String filePath = "C:\\Users\\zby\\Desktop\\custom.txt";
        List<String> lines = FileUtil.readUtf8Lines(filePath);
        System.out.println(JSONUtil.toJsonStr(lines));
    }
}
