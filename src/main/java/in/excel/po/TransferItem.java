package in.excel.po;

import java.util.Random;

public class TransferItem {
    /**
     * 年
     */
    private String year;
    /**
     * 月
     */
    private String month;

    private String day;
    /**
     * 回单编号
     */
    private String transferNum;
    /**
     * 付款人户名
     */
    private String payAccount;
    /**
     * 付款人卡号
     */
    private String payCard;
    /**
     * 付款人开户行
     */
    private String payBank;

    private String toAccount;

    private String toCard;

    private String toBank;

    private Double price;

    /**
     * 业务种类
     */
    private String businessType;
    /**
     * 摘要
     */
    private String tips;

    /**
     * 交易代号
     */
    private String tradeCode;
    /**
     * 支付交易序号
     */
    private String payTradeCode;

    public String getMonth() {
        return month;
    }

    public void setMonth(String month) {
        this.month = month;
    }

    public String getDay() {
        return day;
    }

    public void setDay(String day) {
        this.day = day;
    }

    public String getTransferNum() {
        if (transferNum == null || transferNum.length() == 0) {
            String x = String.format("%04d",new Random().nextInt(9999));
            String y = String.format("%04d",new Random().nextInt(9999));
            transferNum = "0035-" + x + "-" + y + "-1100";
        }
        return transferNum;
    }

    public void setTransferNum(String transferNum) {
        this.transferNum = transferNum;
    }

    public String getPayAccount() {
        return payAccount;
    }

    public void setPayAccount(String payAccount) {
        this.payAccount = payAccount;
    }

    public String getPayCard() {
        return payCard;
    }

    public void setPayCard(String payCard) {
        this.payCard = payCard;
    }

    public String getPayBank() {
        return payBank;
    }

    public void setPayBank(String payBank) {
        this.payBank = payBank;
    }

    public String getToAccount() {
        return toAccount;
    }

    public void setToAccount(String toAccount) {
        this.toAccount = toAccount;
    }

    public String getToCard() {
        return toCard;
    }

    public void setToCard(String toCard) {
        this.toCard = toCard;
    }

    public String getToBank() {
        return toBank;
    }

    public void setToBank(String toBank) {
        this.toBank = toBank;
    }

    public Double getPrice() {
        return price;
    }

    public void setPrice(Double price) {
        this.price = price;
    }

    public String getBusinessType() {
        return businessType;
    }

    public void setBusinessType(String businessType) {
        this.businessType = businessType;
    }

    public String getYear() {
        return year;
    }

    public void setYear(String year) {
        this.year = year;
    }

    public String getTips() {
        return tips;
    }

    public void setTips(String tips) {
        this.tips = tips;
    }

    public String getTradeCode() {
        if (tradeCode == null || tradeCode.length() == 0) {
            tradeCode = String.valueOf((int)((Math.random()*9+1)*10000));
        }
        return tradeCode;
    }

    public void setTradeCode(String tradeCode) {
        this.tradeCode = tradeCode;
    }

    public String getPayTradeCode() {
        if (payTradeCode == null || payTradeCode.length() == 0) {
            payTradeCode = String.valueOf((int)((Math.random()*9+1)*10000000));
        }
        return payTradeCode;
    }

    public void setPayTradeCode(String payTradeCode) {
        this.payTradeCode = payTradeCode;
    }
}
