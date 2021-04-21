package in.excel.po;

import java.util.*;

public class TestItem {
    private static final String EQUALS = "=";
    // 随机的运算符
    private String[] operator = {"+", "-", "x"};
    // 次数对应的分数
    private int[] scores = {0, 2, 5, 10};
    //计算结果
    private int result;
    //第一个运算数
    private final int firstNum;
    //第二个运算数
    private final int secondNum;
    // 运算符
    private final String opt;
    // 分数
    private int score;
    // 答题次数
    private int times;
    // 答题超时时间
    private long timeout;
    // 答题开始时间
    private long start;

    public TestItem() {
        timeout = 10000; // 超时时间
        start = System.currentTimeMillis();
        score = 0; //初始化分数
        times = scores.length - 1;
        opt = operator[new Random().nextInt(operator.length)];
        if ("-".equals(opt)) {
            //2~9
            firstNum = new Random().nextInt(8) + 2;
            //1~firstNum
            secondNum = new Random().nextInt(firstNum - 1) + 1;
        } else {
            //1~9
            firstNum = new Random().nextInt(9) + 1;
            //1~9
            secondNum = new Random().nextInt(9) + 1;
        }
        switch (opt) {
            case "+":
                result = firstNum + secondNum;
                break;
            case "-":
                result = firstNum - secondNum;
                break;
            case "x":
                result = firstNum * secondNum;
                break;
        }
    }

    public boolean isTimeout() {
        return System.currentTimeMillis() - start > timeout;
    }

    public String print() {
        return firstNum + opt + secondNum + EQUALS;
    }

    public boolean ask(int input) {
        if (input == result) {
            score = scores[times];
            System.out.println("恭喜你答对了！本题得分：" + score);
            System.out.println("-------------------------------");
            return true;
        } else {
            times--;
            score = scores[times];
            if (isHasTimes()) {
                System.out.print("答错了，你还有" + times + "次机会！请再次输入：");
            } else {
                System.out.println("三次均答错，本题得分:" + score);
                System.out.println("-------------------------------");
            }
            return false;
        }
    }

    public boolean isHasTimes() {
        return times > 0;
    }

    public int getScore() {
        return score;
    }

    public static void main(String[] args) {
        int count = 10;
        List<TestItem> tests = new ArrayList<>();
        for (int i = 0;i < count;i++) {
            TestItem item = new TestItem();
            tests.add(item);
            System.out.print("第" + (i + 1) + "题：" + item.print());
            boolean isOk = false;
            do {
                try {
                    Scanner scanner = new Scanner(System.in);
                    int input = scanner.nextInt();
                    if (item.isTimeout()) {
                        System.out.println("本题已超时！");
                        break;
                    }
                    isOk = item.ask(input);
                } catch (Exception exception) {
                    System.out.print("请输入整数！再次输入：");
                    continue;
                }
            } while (!isOk && item.isHasTimes());
        }
        int total = tests.stream().mapToInt(TestItem::getScore).sum();
        System.out.println("你的最终得分：" + total);

    }
}
