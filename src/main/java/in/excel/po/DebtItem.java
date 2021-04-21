package in.excel.po;

import java.util.ArrayList;
import java.util.List;

public class DebtItem {
    /**
     * 代码
     */
    private String code;

    /**
     * 客户名称
     */
    private String custom;

    private List<Item> items;

    /**
     * 总金额
     */
    private String totalAccount;

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public String getCustom() {
        return custom;
    }

    public void setCustom(String custom) {
        this.custom = custom;
    }

    public List<Item> getItems() {
        if (items == null) {
            items = new ArrayList<>();
        }
        return items;
    }

    public void setItems(List<Item> items) {
        this.items = items;
    }

    public String getTotalAccount() {
        return totalAccount;
    }

    public void setTotalAccount(String totalAccount) {
        this.totalAccount = totalAccount;
    }

    public void addItem(String name, String unit, String count, String price, String account) {
        Item item = new Item();
        item.setAccount(account);
        item.setCount(count);
        item.setName(name);
        item.setPrice(price);
        item.setUnit(unit);
        getItems().add(item);
    }
}

