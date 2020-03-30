package com.forcewin.excelWriter.vo;

public class RepoVO{
    private int repo_cd;
    private String class_name; 
    private String url_format;

    public int getRepo_cd() {
        return repo_cd;
    }

    public void setRepo_cd(int repo_cd) {
        this.repo_cd = repo_cd;
    }

    public String getClass_name() {
        return class_name;
    }

    public void setClass_name(String class_name) {
        this.class_name = class_name;
    }

    public String getUrl_format() {
        return url_format;
    }

    public void setUrl_format(String url_format) {
        this.url_format = url_format;
    }
    
}