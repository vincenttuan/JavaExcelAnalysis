package com;

import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelBean {
    private String branch;
    private Date date;
    private String reason;

    public String getBranch() {
        return branch;
    }

    public void setBranch(String branch) {
        this.branch = branch;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public String getReason() {
        return reason;
    }

    public void setReason(String reason) {
        this.reason = reason;
    }

    @Override
    public String toString() {
        return "ExcelBean{" + "branch=" + branch + ", date=" + date + ", reason=" + reason + '}';
    }

    
    
    
}
