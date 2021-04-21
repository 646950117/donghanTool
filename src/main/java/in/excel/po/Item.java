package in.excel.po;

public class Item {

    /**
     * 产品名称
     */
    private String name;

    /**
     * 单位
     */
    private String unit;

    /**
     * 数量
     */
    private String count;

    /**
     * 单价
     */
    private String price;

    /**
     * 金额
     */
    private String account;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getUnit() {
        return unit;
    }

    public void setUnit(String unit) {
        this.unit = unit;
    }

    public String getCount() {
        return count;
    }

    public void setCount(String count) {
        this.count = count;
    }

    public String getPrice() {
        return price;
    }

    public void setPrice(String price) {
        this.price = price;
    }

    public String getAccount() {
        return account;
    }

    public void setAccount(String account) {
        this.account = account;
    }

}
